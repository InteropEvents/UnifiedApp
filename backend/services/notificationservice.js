const { WssSocketService } = require('./wsssocketservice');

class NotificationService {
    async processNotification(notification) {
        WssSocketService.broadcast(JSON.stringify(notification));
        console.log('POST /processNotification');
    }

    async handleNotifications(notifications = []) {
        for (const notification of notifications) {
            await this.processNotification(notification);
        }
    }
}

module.exports = { NotificationService: new NotificationService() };
