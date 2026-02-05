# WinRM Bridge

Concurrent PowerShell session manager with real-time WebSocket streaming.

## Features

- **Multiple concurrent sessions** - Local or remote WinRM connections
- **WebSocket streaming** - Real-time command output
- **Timeout protection** - Commands auto-cancel after 30s (configurable)
- **Cancel/Interrupt** - Stop stuck commands instantly
- **Web UI** - Mobile-friendly interface
- **Hive mesh integration** - Auto-registers with orchestrator

## Quick Start

```bash
npm install
node server.js
```

Open http://localhost:8775

## API Reference

### Sessions

| Method | Endpoint | Description |
|--------|----------|-------------|
| `POST` | `/session/create` | Create new session |
| `GET` | `/sessions` | List all sessions |
| `GET` | `/session/:id` | Get session info |
| `PATCH` | `/session/:id` | Update session label |
| `DELETE` | `/session/:id` | Close session |
| `DELETE` | `/sessions/all` | Close all sessions |

### Command Execution

| Method | Endpoint | Description |
|--------|----------|-------------|
| `POST` | `/session/:id/exec` | Execute command (with timeout) |
| `POST` | `/session/:id/cancel` | Cancel running command |
| `POST` | `/sessions/exec-all` | Broadcast to all sessions |
| `POST` | `/sessions/cancel-all` | Cancel all running commands |

### WebSocket

Connect: `ws://localhost:8775?session=<sessionId>`

**Send:**
```json
{"type": "exec", "command": "Get-Date", "timeout": 30000}
{"type": "interrupt"}
```

**Receive:**
```json
{"type": "output", "data": "..."}
{"type": "result", "success": true, "output": "...", "executionTime": 150}
{"type": "error", "message": "..."}
{"type": "interrupted"}
```

## Examples

### Create Local Session
```bash
curl -X POST http://localhost:8775/session/create \
  -H "Content-Type: application/json" \
  -d '{"label": "Local PS"}'
```

### Create Remote Session
```bash
curl -X POST http://localhost:8775/session/create \
  -H "Content-Type: application/json" \
  -d '{
    "computer": "192.168.1.42",
    "label": "Remote-PC",
    "credentials": {
      "username": "DOMAIN\\user",
      "password": "password"
    }
  }'
```

### Execute with Timeout
```bash
curl -X POST http://localhost:8775/session/<id>/exec \
  -H "Content-Type: application/json" \
  -d '{"command": "Get-Process", "timeout": 5000}'
```

### Cancel Running Command
```bash
curl -X POST http://localhost:8775/session/<id>/cancel
```

## Configuration

| Variable | Default | Description |
|----------|---------|-------------|
| `PORT` | 8775 | HTTP/WebSocket port |
| `DEFAULT_TIMEOUT` | 30000 | Command timeout (ms) |
| `HIVE_MESH` | localhost:8750 | Orchestrator address |

## Installation as Service

Run `INSTALL-IMMORTAL.bat` to install as Windows service (auto-starts with Windows).

## Port

**8775** - HTTP API + WebSocket
