const winston = require('winston');
const path = require('path');
const fs = require('fs');

// Ensure logs directory exists
const logsDir = path.join(__dirname, '../../logs');
if (!fs.existsSync(logsDir)) {
    fs.mkdirSync(logsDir, { recursive: true });
}

// Create the logger
const logger = winston.createLogger({
    level: 'info',
    format: winston.format.combine(
        winston.format.timestamp(),
        winston.format.json()
    ),
    transports: [
        // Write to all logs with level 'info' and below to `combined.log`
        new winston.transports.File({ 
            filename: path.join(logsDir, 'combined.log')
        }),
        // Write all logs error (and below) to `error.log`
        new winston.transports.File({ 
            filename: path.join(logsDir, 'error.log'), 
            level: 'error' 
        }),
        // Write to console with custom format
        new winston.transports.Console({
            format: winston.format.combine(
                winston.format.colorize(),
                winston.format.simple(),
                winston.format.printf(({ level, message, timestamp, ...metadata }) => {
                    let msg = `${timestamp} ${level}: ${message}`;
                    if (Object.keys(metadata).length > 0) {
                        msg += JSON.stringify(metadata);
                    }
                    return msg;
                })
            )
        })
    ]
});

// Add convenience methods that maintain console.log compatibility
const enhancedLogger = {
    info: (message, meta = {}) => {
        console.log(message); // Maintain existing console.log behavior
        logger.info(message, meta);
    },
    error: (message, error = null, meta = {}) => {
        console.error(message); // Maintain existing console.error behavior
        if (error) {
            console.error(error);
            logger.error(message, { ...meta, error: error.message, stack: error.stack });
        } else {
            logger.error(message, meta);
        }
    },
    debug: (message, meta = {}) => {
        console.log(message); // Maintain existing console.log behavior
        logger.debug(message, meta);
    },
    warn: (message, meta = {}) => {
        console.warn(message); // Maintain existing console.warn behavior
        logger.warn(message, meta);
    }
};

module.exports = enhancedLogger;