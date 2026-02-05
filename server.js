const express = require('express');
const { PowerShell } = require('node-powershell');
const axios = require('axios');
const { v4: uuidv4 } = require('uuid');
const { spawn } = require('child_process');

const PORT = 8775;
const SERVICE_NAME = 'WinRM-Bridge';
const HIVE_MESH = 'http://localhost:8750';

const http = require('http');
const path = require('path');
const WebSocket = require('ws');

// Shell configurations for local terminals
const SHELLS = {
  powershell: { path: 'powershell.exe', args: ['-NoLogo', '-NoExit'], name: 'PowerShell' },
  cmd: { path: 'cmd.exe', args: ['/K'], name: 'CMD' },
  bash: { path: 'C:\\Program Files\\Git\\usr\\bin\\bash.exe', args: ['--login', '-i'], name: 'Git Bash' }
};

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Create HTTP server for both Express and WebSocket
const server = http.createServer(app);
const wss = new WebSocket.Server({ server });

// Active sessions - fully concurrent
const sessions = new Map();

// Running commands - track for cancellation
const runningCommands = new Map();

// Default timeout (30 seconds)
const DEFAULT_TIMEOUT = 30000;

// WebSocket connection handling
wss.on('connection', (ws, req) => {
  const url = new URL(req.url, `http://${req.headers.host}`);
  const sessionId = url.searchParams.get('session');

  if (!sessionId || !sessions.has(sessionId)) {
    ws.send(JSON.stringify({ type: 'error', message: 'Invalid session ID' }));
    ws.close();
    return;
  }

  const session = sessions.get(sessionId);
  session.clients.add(ws);
  console.log(`⚡ WebSocket client connected to session ${session.label}`);

  // Send buffer (recent output)
  if (session.buffer.length > 0) {
    ws.send(JSON.stringify({ type: 'output', data: session.buffer }));
  }

  ws.on('message', async (msg) => {
    try {
      const data = JSON.parse(msg);

      if (data.type === 'input') {
        // Direct input for local terminals
        if (session.type === 'local' && session.process && !session.process.killed) {
          session.process.stdin.write(data.data);
          session.lastUsed = Date.now();
        } else {
          ws.send(JSON.stringify({ type: 'error', message: 'No input stream available' }));
        }
      } else if (data.type === 'exec') {
        // Execute command (works for both local and WinRM)
        const command = data.command;
        const timeout = data.timeout || DEFAULT_TIMEOUT;

        if (session.type === 'local') {
          // For local terminals, just send as input
          if (session.process && !session.process.killed) {
            session.process.stdin.write(command + '\r\n');
            session.lastUsed = Date.now();
            session.commandCount++;
          }
        } else {
          // For WinRM, use executeWithStreaming
          ws.send(JSON.stringify({ type: 'status', status: 'executing', command }));
          try {
            const result = await executeWithStreaming(sessionId, command, timeout);
            ws.send(JSON.stringify({ type: 'result', ...result }));
          } catch (err) {
            ws.send(JSON.stringify({ type: 'error', message: err.message }));
          }
        }
      } else if (data.type === 'interrupt') {
        // Send Ctrl+C for local, cancel for WinRM
        if (session.type === 'local' && session.process && !session.process.killed) {
          session.process.stdin.write('\x03');
        }
        cancelRunningCommand(sessionId);
        ws.send(JSON.stringify({ type: 'interrupted' }));
      }
    } catch (e) {
      console.error('WebSocket message error:', e);
    }
  });

  ws.on('close', () => {
    session.clients.delete(ws);
    console.log(`WebSocket client disconnected from session ${session.label}`);
  });
});

// Execute command with streaming output to WebSocket clients
async function executeWithStreaming(sessionId, command, timeout) {
  const session = sessions.get(sessionId);
  if (!session) throw new Error('Session not found');

  if (runningCommands.has(sessionId)) {
    throw new Error('Command already running');
  }

  const commandId = uuidv4();
  let cancelled = false;
  let timeoutId = null;

  runningCommands.set(sessionId, {
    commandId,
    command,
    startTime: Date.now(),
    cancel: () => { cancelled = true; }
  });

  const broadcast = (data) => {
    const msg = JSON.stringify(data);
    session.clients.forEach(ws => {
      if (ws.readyState === WebSocket.OPEN) {
        ws.send(msg);
      }
    });
  };

  try {
    console.log(`⚡ [WS] Executing in ${session.label}: ${command}`);
    const startTime = Date.now();

    // Broadcast that we're starting
    broadcast({ type: 'output', data: `❯ ${command}\n` });
    session.buffer += `❯ ${command}\n`;

    const result = await Promise.race([
      session.ps.invoke(command),
      new Promise((_, reject) => {
        timeoutId = setTimeout(() => reject(new Error('Command timed out')), timeout);
      }),
      new Promise((_, reject) => {
        const check = setInterval(() => {
          if (cancelled) {
            clearInterval(check);
            reject(new Error('Command cancelled'));
          }
        }, 100);
        runningCommands.get(sessionId).cancelInterval = check;
      })
    ]);

    clearTimeout(timeoutId);
    const executionTime = Date.now() - startTime;
    const output = result.raw || '(no output)';

    // Broadcast output
    broadcast({ type: 'output', data: output + '\n' });
    session.buffer += output + '\n';
    if (session.buffer.length > 100000) session.buffer = session.buffer.slice(-50000);

    session.lastUsed = Date.now();
    session.commandCount++;

    return { success: true, output, executionTime, commandCount: session.commandCount };
  } catch (error) {
    clearTimeout(timeoutId);
    const errMsg = `[Error: ${error.message}]\n`;
    broadcast({ type: 'output', data: errMsg });
    session.buffer += errMsg;
    return { success: false, error: error.message, timedOut: error.message.includes('timed out') };
  } finally {
    const cmd = runningCommands.get(sessionId);
    if (cmd?.cancelInterval) clearInterval(cmd.cancelInterval);
    runningCommands.delete(sessionId);
  }
}

// Cancel running command helper
function cancelRunningCommand(sessionId) {
  const running = runningCommands.get(sessionId);
  if (running) {
    running.cancel();
    console.log(`⚠ Cancelled command in session ${sessionId}`);
  }
}

// Create new session (local terminal or remote WinRM)
app.post('/session/create', async (req, res) => {
  const { type = 'local', shell = 'powershell', computer = 'localhost', credentials, label, cwd } = req.body;
  const sessionId = uuidv4();

  try {
    let session;

    if (type === 'local') {
      // Spawn local terminal process
      const shellConfig = SHELLS[shell] || SHELLS.powershell;
      const workDir = cwd || process.env.USERPROFILE || 'C:\\';

      const proc = spawn(shellConfig.path, shellConfig.args, {
        cwd: workDir,
        env: { ...process.env, TERM: 'xterm-256color' },
        stdio: ['pipe', 'pipe', 'pipe'],
        shell: false
      });

      session = {
        type: 'local',
        process: proc,
        shell,
        computer: 'localhost',
        label: label || `${shellConfig.name} ${sessions.size + 1}`,
        created: Date.now(),
        lastUsed: Date.now(),
        commandCount: 0,
        clients: new Set(),
        buffer: ''
      };

      // Stream stdout/stderr to WebSocket clients
      const broadcast = (data) => {
        const text = data.toString();
        session.buffer += text;
        if (session.buffer.length > 100000) session.buffer = session.buffer.slice(-50000);
        const msg = JSON.stringify({ type: 'output', data: text });
        session.clients.forEach(ws => {
          if (ws.readyState === WebSocket.OPEN) ws.send(msg);
        });
      };

      proc.stdout.on('data', broadcast);
      proc.stderr.on('data', broadcast);
      proc.on('exit', (code) => {
        const msg = `\r\n[Process exited with code ${code}]\r\n`;
        session.buffer += msg;
        session.clients.forEach(ws => {
          if (ws.readyState === WebSocket.OPEN) {
            ws.send(JSON.stringify({ type: 'output', data: msg }));
            ws.send(JSON.stringify({ type: 'exit', code }));
          }
        });
      });
      proc.on('error', (err) => {
        console.error(`Terminal error: ${err.message}`);
      });

      // Initial prompt
      setTimeout(() => {
        if (!proc.killed) {
          if (shell === 'powershell') proc.stdin.write('Write-Output "[Terminal Ready] $(Get-Location)"\r\n');
          else if (shell === 'cmd') proc.stdin.write('echo [Terminal Ready] & cd\r\n');
          else proc.stdin.write('echo "[Terminal Ready] $(pwd)"\n');
        }
      }, 300);

    } else {
      // WinRM session via node-powershell
      const ps = new PowerShell({
        executionPolicy: 'Bypass',
        noProfile: true
      });

      if (computer !== 'localhost' && credentials) {
        const credCmd = `$cred = New-Object System.Management.Automation.PSCredential('${credentials.username}', (ConvertTo-SecureString '${credentials.password}' -AsPlainText -Force))`;
        await ps.invoke(credCmd);
        await ps.invoke(`$session = New-PSSession -ComputerName ${computer} -Credential $cred`);
      }

      session = {
        type: 'winrm',
        ps,
        computer,
        label: label || `WinRM-${computer}`,
        created: Date.now(),
        lastUsed: Date.now(),
        commandCount: 0,
        clients: new Set(),
        buffer: ''
      };
    }

    sessions.set(sessionId, session);
    console.log(`✓ Created ${session.type} session: ${session.label} (${sessionId})`);

    res.json({
      success: true,
      sessionId,
      type: session.type,
      computer: session.computer,
      shell: session.shell,
      label: session.label,
      message: 'Session created'
    });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// Send command to session (non-blocking for other sessions)
app.post('/session/:sessionId/exec', async (req, res) => {
  const { sessionId } = req.params;
  const { command, timeout = DEFAULT_TIMEOUT } = req.body;

  const session = sessions.get(sessionId);
  if (!session) {
    return res.status(404).json({ success: false, error: 'Session not found' });
  }

  // Check if command already running
  if (runningCommands.has(sessionId)) {
    return res.status(409).json({ success: false, error: 'Command already running. Cancel it first.' });
  }

  const commandId = uuidv4();
  let cancelled = false;
  let timeoutId = null;

  // Track this command
  runningCommands.set(sessionId, {
    commandId,
    command,
    startTime: Date.now(),
    cancel: () => { cancelled = true; }
  });

  try {
    console.log(`⚡ Executing in ${session.label}: ${command}`);
    const startTime = Date.now();

    // Race between command execution and timeout
    const result = await Promise.race([
      session.ps.invoke(command),
      new Promise((_, reject) => {
        timeoutId = setTimeout(() => {
          reject(new Error(`Command timed out after ${timeout}ms`));
        }, timeout);
      }),
      new Promise((_, reject) => {
        const checkCancelled = setInterval(() => {
          if (cancelled) {
            clearInterval(checkCancelled);
            reject(new Error('Command cancelled'));
          }
        }, 100);
        // Store interval for cleanup
        runningCommands.get(sessionId).cancelInterval = checkCancelled;
      })
    ]);

    clearTimeout(timeoutId);
    const executionTime = Date.now() - startTime;

    session.lastUsed = Date.now();
    session.commandCount++;

    console.log(`✓ Command completed in ${executionTime}ms`);

    res.json({
      success: true,
      output: result.raw,
      sessionId,
      label: session.label,
      executionTime,
      commandCount: session.commandCount
    });
  } catch (error) {
    clearTimeout(timeoutId);
    console.log(`✗ Command failed: ${error.message}`);
    res.status(500).json({
      success: false,
      error: error.message,
      sessionId,
      label: session.label,
      timedOut: error.message.includes('timed out'),
      cancelled: error.message.includes('cancelled')
    });
  } finally {
    // Cleanup
    const cmd = runningCommands.get(sessionId);
    if (cmd?.cancelInterval) clearInterval(cmd.cancelInterval);
    runningCommands.delete(sessionId);
  }
});

// Cancel running command
app.post('/session/:sessionId/cancel', async (req, res) => {
  const { sessionId } = req.params;

  const running = runningCommands.get(sessionId);
  if (!running) {
    return res.status(404).json({ success: false, error: 'No command running' });
  }

  console.log(`⚠ Cancelling command in session ${sessionId}`);
  running.cancel();

  // Also try to kill the PowerShell process and recreate
  const session = sessions.get(sessionId);
  if (session) {
    try {
      await session.ps.dispose();
      session.ps = new PowerShell({
        executionPolicy: 'Bypass',
        noProfile: true
      });
      console.log(`✓ Recreated PowerShell instance for ${session.label}`);
    } catch (e) {
      console.log(`⚠ Failed to recreate PowerShell: ${e.message}`);
    }
  }

  res.json({ success: true, message: 'Cancel signal sent', commandId: running.commandId });
});

// Execute command across multiple sessions concurrently
app.post('/sessions/exec-all', async (req, res) => {
  const { command, sessionIds, timeout } = req.body;

  const targetSessions = sessionIds || Array.from(sessions.keys());
  console.log(`⚡ Executing across ${targetSessions.length} sessions concurrently...`);

  const results = await Promise.all(
    targetSessions.map(async (sessionId) => {
      try {
        const response = await axios.post(`http://localhost:${PORT}/session/${sessionId}/exec`, { command, timeout });
        return { sessionId, ...response.data };
      } catch (error) {
        return { sessionId, success: false, error: error.response?.data?.error || error.message };
      }
    })
  );

  res.json({ results, count: results.length });
});

// Cancel all running commands
app.post('/sessions/cancel-all', async (req, res) => {
  const cancelled = [];
  for (const sessionId of runningCommands.keys()) {
    try {
      await axios.post(`http://localhost:${PORT}/session/${sessionId}/cancel`);
      cancelled.push(sessionId);
    } catch (error) {
      console.log(`Failed to cancel ${sessionId}: ${error.message}`);
    }
  }
  res.json({ success: true, cancelled, count: cancelled.length });
});

// Get session info
app.get('/session/:sessionId', (req, res) => {
  const { sessionId } = req.params;
  const session = sessions.get(sessionId);

  if (!session) {
    return res.status(404).json({ success: false, error: 'Session not found' });
  }

  const running = runningCommands.get(sessionId);
  res.json({
    sessionId,
    label: session.label,
    computer: session.computer,
    created: session.created,
    lastUsed: session.lastUsed,
    uptime: Date.now() - session.created,
    commandCount: session.commandCount,
    running: running ? { command: running.command, elapsed: Date.now() - running.startTime } : null
  });
});

// Update session label
app.patch('/session/:sessionId', (req, res) => {
  const { sessionId } = req.params;
  const { label } = req.body;
  
  const session = sessions.get(sessionId);
  if (!session) {
    return res.status(404).json({ success: false, error: 'Session not found' });
  }

  session.label = label;
  console.log(`✓ Updated session label to: ${label}`);
  res.json({ success: true, sessionId, label });
});

// Close session
app.delete('/session/:sessionId', async (req, res) => {
  const { sessionId } = req.params;
  const session = sessions.get(sessionId);

  if (!session) {
    return res.status(404).json({ success: false, error: 'Session not found' });
  }

  try {
    if (session.type === 'local' && session.process) {
      session.process.kill();
    } else if (session.ps) {
      await session.ps.dispose();
    }
    sessions.delete(sessionId);
    console.log(`✓ Closed session: ${session.label}`);
    res.json({ success: true, message: 'Session closed', sessionId });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// Close all sessions
app.delete('/sessions/all', async (req, res) => {
  const closed = [];
  for (const [id, session] of sessions.entries()) {
    try {
      if (session.type === 'local' && session.process) {
        session.process.kill();
      } else if (session.ps) {
        await session.ps.dispose();
      }
      sessions.delete(id);
      closed.push(id);
    } catch (error) {
      console.error(`Failed to close session ${id}:`, error.message);
    }
  }
  console.log(`✓ Closed ${closed.length} sessions`);
  res.json({ success: true, closed, count: closed.length });
});

// List all sessions
app.get('/sessions', (req, res) => {
  const sessionList = Array.from(sessions.entries()).map(([id, session]) => {
    const running = runningCommands.get(id);
    return {
      sessionId: id,
      type: session.type || 'winrm',
      shell: session.shell,
      label: session.label,
      computer: session.computer,
      created: session.created,
      lastUsed: session.lastUsed,
      uptime: Date.now() - session.created,
      commandCount: session.commandCount,
      running: running ? { command: running.command, elapsed: Date.now() - running.startTime } : null,
      alive: session.type === 'local' ? (session.process && !session.process.killed) : true
    };
  });

  res.json({
    sessions: sessionList,
    count: sessionList.length,
    totalCommands: sessionList.reduce((sum, s) => sum + s.commandCount, 0),
    runningCount: runningCommands.size
  });
});

// Health check
app.get('/health', (req, res) => {
  res.json({
    service: SERVICE_NAME,
    status: 'healthy',
    port: PORT,
    uptime: process.uptime(),
    activeSessions: sessions.size
  });
});

// Clean up stale sessions (30 min idle)
setInterval(() => {
  const now = Date.now();
  for (const [id, session] of sessions.entries()) {
    if (now - session.lastUsed > 1800000) {
      session.ps.dispose();
      sessions.delete(id);
      console.log(`Cleaned up stale session: ${session.label} (${id})`);
    }
  }
}, 60000);

// Register with Hive mesh via WebSocket
function registerWithHive() {
  try {
    const ws = new WebSocket(`ws://localhost:8750`);
    ws.on('open', () => {
      ws.send(JSON.stringify({
        type: 'agent:register',
        data: {
          id: `winrm-bridge-${require('os').hostname()}`,
          type: 'winrm-bridge',
          status: 'online',
          port: PORT,
          capabilities: ['winrm', 'powershell', 'concurrent-sessions'],
          ui: `http://localhost:${PORT}`
        }
      }));
      console.log(`✓ Registered with Hive mesh via WebSocket`);
    });
    ws.on('error', () => {
      console.log(`⚠ Hive mesh not available, running standalone`);
    });
    ws.on('close', () => {
      setTimeout(registerWithHive, 30000);
    });
  } catch (e) {
    console.log(`⚠ Hive mesh not available, running standalone`);
  }
}

server.listen(PORT, async () => {
  console.log(`╔════════════════════════════════════════╗`);
  console.log(`║  ${SERVICE_NAME} Running               ║`);
  console.log(`╚════════════════════════════════════════╝`);
  console.log(`Port: ${PORT}`);
  console.log(`HTTP + WebSocket ready`);
  console.log(`WebSocket: ws://localhost:${PORT}?session=<id>`);
  console.log(``);
  await registerWithHive();
});
