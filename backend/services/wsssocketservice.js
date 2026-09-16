const WebSocket = require('ws');

let wss; // Declare wss outside to make it accessible within the class

class WssSocketService {
    // Initialize the WebSocket server with the provided HTTP server instance
    static initialize(server) {
        wss = new WebSocket.Server({ server });
        wss.on('connection', (ws) => {
            console.log('Client connected');
            ws.on('close', () => {
                console.log('Client disconnected');
            });
        });
    }

    // Broadcast a message to all connected clients
    static broadcast(message) {
        if (!wss) {
            throw new Error('WebSocket server is not initialized.');
        }
        wss.clients.forEach(client => {
            if (client.readyState === WebSocket.OPEN) {
                client.send(message);
            }
        });
    }
}

module.exports = { WssSocketService };