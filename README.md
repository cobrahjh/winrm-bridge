# WinRM Bridge - Concurrent Session PowerShell Service

Web-based REST API for persistent PowerShell/WinRM sessions with full concurrency support.

## Installation

```batch
cd C:\DevClaude\Hivemind\services\winrm-bridge
INSTALL-IMMORTAL.bat
```

Service runs on port **8775** and auto-starts with Windows.

## API Endpoints

### Create Session
```bash
POST /session/create
{
  "computer": "localhost",           # or "192.168.1.97"
  "label": "My-Session",            # optional friendly name
  "credentials": {                   # optional for remote
    "username": "DOMAIN\\user",
    "password": "password"
  }
}
```

### Execute Command
```bash
POST /session/:sessionId/exec
{"command": "Get-Process | Select-Object -First 5"}
```

### Execute Across All Sessions
```bash
POST /sessions/exec-all
{"command": "hostname"}
```
