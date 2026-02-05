const express = require('express');
const { PowerShell } = require('node-powershell');
const axios = require('axios');
const { v4: uuidv4 } = require('uuid');

const PORT = 8775;
const SERVICE_NAME = 'WinRM-Bridge';
const HIVE_MESH = 'http://localhost:8750';

const path = require('path');
const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Active sessions - fully concurrent
const sessions = new Map();

// Running commands - track for cancellation
const runningCommands = new Map();

// Default timeout (30 seconds)
const DEFAULT_TIMEOUT = 30000;

// Create new session
app.post('/session/create', async (req, res) => {
  const { computer = 'localhost', credentials, label } = req.body;
  const sessionId = uuidv4();

  try {
    const ps = new PowerShell({
      executionPolicy: 'Bypass',
      noProfile: true
    });

    // Initialize remote session if needed
    if (computer !== 'localhost' && credentials) {
      const credCmd = `$cred = New-Object System.Management.Automation.PSCredential('${credentials.username}', (ConvertTo-SecureString '${credentials.password}' -AsPlainText -Force))`;
      await ps.invoke(credCmd);
      await ps.invoke(`$session = New-PSSession -ComputerName ${computer} -Credential $cred`);
    }

    sessions.set(sessionId, {
      ps,
      computer,
      label: label || `Session-${sessions.size + 1}`,
      created: Date.now(),
      lastUsed: Date.now(),
      commandCount: 0
    });

    console.log(`✓ Created session: ${sessions.get(sessionId).label} (${sessionId})`);

    res.json({ 
      success: true, 
      sessionId,
      computer,
      label: sessions.get(sessionId).label,
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
    await session.ps.dispose();
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
      await session.ps.dispose();
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
      label: session.label,
      computer: session.computer,
      created: session.created,
      lastUsed: session.lastUsed,
      uptime: Date.now() - session.created,
      commandCount: session.commandCount,
      running: running ? { command: running.command, elapsed: Date.now() - running.startTime } : null
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
const WebSocket = require('ws');
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

app.listen(PORT, async () => {
  console.log(`╔════════════════════════════════════════╗`);
  console.log(`║  ${SERVICE_NAME} Running               ║`);
  console.log(`╚════════════════════════════════════════╝`);
  console.log(`Port: ${PORT}`);
  console.log(`Ready for concurrent session management`);
  console.log(``);
  await registerWithHive();
});
