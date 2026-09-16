import { NotificationClientInstance } from './NotificationClient';

class WebSocketClient {
  private static instance: WebSocketClient;
  private socket: WebSocket | null = null;

  private constructor() {}

  static getInstance(): WebSocketClient {
    if (!WebSocketClient.instance) {
      WebSocketClient.instance = new WebSocketClient();
    }
    return WebSocketClient.instance;
  }

  connect() {
    if (this.socket) {
      console.warn('WebSocket is already connected.');
      return;
    }

    this.socket = new WebSocket(`wss://${process.env.REACT_APP_DOMAIN || 'localhost:3001'}`);

    this.socket.onopen = () => console.log('WebSocket connection established.');
    this.socket.onmessage = this.handleMessage.bind(this);
    this.socket.onclose = () => {
      console.log('WebSocket connection closed.');
      this.socket = null;
    };
    this.socket.onerror = (error) => console.error('WebSocket error:', error);
  }

  private async handleMessage(event: MessageEvent) {
    try {
        NotificationClientInstance.handleNotification(event);
    } catch (error) {
      console.error('Failed to process WebSocket message:', error);
    }
  }

  disconnect() {
    this.socket?.close();
    this.socket = null;
  }
}

export const WebSocketClientInstance = WebSocketClient.getInstance();
