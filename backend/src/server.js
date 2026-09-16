const express = require("express");
const http = require('http');
const bodyParser = require("body-parser");
const { getAppLevelToken } = require("../services/tokenService");
const { NotificationService } = require("../services/notificationservice");
const { WssSocketService } = require('../services/wsssocketservice');

class Server {
    constructor() {
        this.app = express();
        this.server = http.createServer(this.app);
        this.PORT = process.env.PORT || 3001;
    }

    initializeMiddlewares() {
        this.app.use(bodyParser.json());
    }

    initializeServices() {
        WssSocketService.initialize(this.server);
    }

    initializeRoutes() {
        this.app.post("/resourceNotifications", this.handleResourceNotifications.bind(this));
        this.app.post("/emailNotification", this.handlEmailNotification.bind(this));
        this.app.get("/getAppToken", this.handleGetAppToken.bind(this));
        this.app.post("/teamsChannelFiles", this.teamsChannelFilesfication.bind(this));
    }

    async handleResourceNotifications(req, res) {
        console.log("Received resource notifications:", req.body);
        if (req.query.validationToken) {
            res.set('Content-Type', 'text/plain');
            res.status(200).send(req.query.validationToken);
        } else {
            try {
                const notifications = req.body.value;
                await NotificationService.handleNotifications(notifications);
                res.status(200).send("Notification received");
            } catch (error) {
                console.error("Error processing notifications:", error);
                res.status(500).send("Failed to process notifications");
            }
        }
    }

    async handlEmailNotification(req, res) {
        if (req.query.validationToken) {
            res.set('Content-Type', 'text/plain');
            res.status(200).send(req.query.validationToken);
        } else {
            try {
                const notifications = req.body.value;
                await NotificationService.handleNotifications(notifications);
                res.status(200).send("Notification received");
            } catch (error) {
                console.error("Error processing notifications:", error);
                res.status(500).send("Failed to process notifications");
            }
        }
    }

    async teamsChannelFilesfication(req, res) {
        if (req.query.validationToken) {
            res.set('Content-Type', 'text/plain');
            res.status(200).send(req.query.validationToken);
        } else {
            try {
                const notifications = req.body.value;
                await NotificationService.handleNotifications(notifications);
                res.status(200).send("Notification received");
            } catch (error) {
                console.error("Error processing notifications:", error);
                res.status(500).send("Failed to process notifications");
            }
        }
    }

    async handleGetAppToken(req, res) {
        try {
            const token = await getAppLevelToken();
            res.json({ token });
        } catch (error) {
            res.status(500).send("Failed to acquire token");
        }
    }

    start() {
        this.initializeMiddlewares();
        this.initializeServices();
        this.initializeRoutes();

        this.server.listen(this.PORT, () => {
            console.log(`Server is running on port ${this.PORT}`);
        });
    }
}

// Instantiate and start the server
new Server().start();
