# UnifiedApp notification backend

The backend receives Microsoft Graph webhook notifications and broadcasts them to connected frontend clients over WebSocket.

## Prerequisites

- Node.js and npm
- A Microsoft Entra app with the required Microsoft Graph permissions
- Microsoft Dev Tunnels CLI

Redis is not required. Notifications are broadcast directly when the webhook receives them.

## Initial setup

1. Copy `env.sample` to `.env`.
2. Configure `CLIENT_ID`, `CLIENT_SECRET`, `TENANT_ID`, and `SCOPE` in `.env`.
3. Install dependencies:

```powershell
cd .\UnifiedApp\backend
npm ci
```

4. Set the Dev Tunnels executable used on this machine:

```powershell
$devtunnel = "$env:LOCALAPPDATA\Programs\DevTunnel\devtunnel.exe"
```

5. Sign in and create the persistent tunnel once:

```powershell
& $devtunnel user login --entra --use-integrated-windows-auth
& $devtunnel create taipeidevdays2026 --allow-anonymous
& $devtunnel port create taipeidevdays2026.use -p 3001 --protocol http
```

## Start the integrated application

Use three terminals and keep all three processes running.

### Terminal 1: backend

```powershell
cd .\UnifiedApp\backend
npm start
```

Expected output:

```text
Server is running on port 3001
```

### Terminal 2: public tunnel

```powershell
$devtunnel = "$env:LOCALAPPDATA\Programs\DevTunnel\devtunnel.exe"
& $devtunnel host taipeidevdays2026.use
```

Copy the hostname from the displayed `Connect via browser` URL. Do not include `https://`, `wss://`, a path, or a trailing slash.

Set that hostname in `frontend/.env`:

```dotenv
REACT_APP_DOMAIN="<tunnel-hostname>"
```

The frontend uses this value for both:

- `https://<tunnel-hostname>/resourceNotifications`
- `wss://<tunnel-hostname>`

### Terminal 3: frontend

Restart the frontend whenever `frontend/.env` changes:

```powershell
cd .\UnifiedApp\frontend
yarn start
```

Open `http://localhost:3000` and sign in. The frontend creates the Microsoft Graph presence subscription after sign-in.

## Verify integration

When the integration is working, the backend terminal shows:

```text
Client connected
Received resource notifications: ...
POST /processNotification
```

The tunnel must remain running for Microsoft Graph webhook delivery and frontend WebSocket updates. The current persistent tunnel expires after 30 days; recreate it and update `REACT_APP_DOMAIN` when its URL changes.