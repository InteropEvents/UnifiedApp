import PQueue from "p-queue";

class NotificationClient {
    private static instance: NotificationClient;
    private notificationQueue: PQueue;

    private constructor() {
        // Private constructor to prevent direct instantiation
        this.notificationQueue = new PQueue({ concurrency: 1 });
    }

    static getInstance(): NotificationClient {
        if (!NotificationClient.instance) {
            NotificationClient.instance = new NotificationClient();
        }
        return NotificationClient.instance;
    }

    async processNotification(notification: any): Promise<void> {
        // Logic to process the notification
        console.log("Processing notification:", notification);
    }

    handleNotification(notification: any): void {
        // Add the notification to the queue
        this.notificationQueue.add(() => this.processNotification(notification));
        console.log("Notification added to queue:", notification);
    }
}

export const NotificationClientInstance = NotificationClient.getInstance();
