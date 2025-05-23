// logger.js (opsional, modular)
const winston = require("winston");

const logger = winston.createLogger({
  level: "debug", // Atur level minimum log
  format: winston.format.combine(
    winston.format.timestamp(),
    winston.format.printf(({ timestamp, level, message }) => {
      return `[${timestamp}] [${level.toUpperCase()}]: ${message}`;
    })
  ),
  transports: [new winston.transports.Console(), new winston.transports.File({ filename: "debug.log" })],
});

module.exports = logger;
