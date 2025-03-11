const path = require('path');
const fs = require('fs');
const fsPromises = fs.promises;
const logger = require('./logger');

// Ensure screenshots directory exists
const screenshotsDir = path.join(__dirname, '../../screenshots');
if (!fs.existsSync(screenshotsDir)) {
    fs.mkdirSync(screenshotsDir, { recursive: true });
}

class ErrorHandler {
    /**
     * Captures the current state of the page when an error occurs.
     * It takes a screenshot, saves the HTML, and collects any pending console messages.
     *
     * @param {object} page - The Playwright page instance.
     * @param {Error} error - The error that occurred.
     * @param {string} [context=''] - Additional context for the error.
     * @returns {Promise<object|null>} An object with paths to the captured files or null on failure.
     */
    static async captureErrorState(page, error, context = '') {
        try {
            const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
            const contextPart = context ? `${context}-` : '';
            const baseFileName = `error-${contextPart}${timestamp}`;

            // Capture screenshot
            const screenshotPath = path.join(screenshotsDir, `${baseFileName}.png`);
            await page.screenshot({ 
                path: screenshotPath,
                fullPage: true 
            });

            // Capture HTML of the page
            const htmlPath = path.join(screenshotsDir, `${baseFileName}.html`);
            const html = await page.evaluate(() => document.documentElement.outerHTML);
            await fsPromises.writeFile(htmlPath, html);

            // Capture console logs:
            // Attach a temporary listener, wait briefly, then remove it.
            const consoleMessages = [];
            const consoleListener = msg => consoleMessages.push(`${msg.type()}: ${msg.text()}`);
            page.on('console', consoleListener);
            await new Promise(resolve => setTimeout(resolve, 100)); // Wait 100ms to catch any messages
            page.removeListener('console', consoleListener);
            const consolePath = path.join(screenshotsDir, `${baseFileName}-console.log`);
            await fsPromises.writeFile(consolePath, consoleMessages.join('\n'));

            logger.info('Error state captured', {
                screenshot: screenshotPath,
                html: htmlPath,
                console: consolePath,
                error: error.message
            });

            return {
                screenshot: screenshotPath,
                html: htmlPath,
                console: consolePath
            };
        } catch (captureError) {
            logger.error('Failed to capture error state:', captureError);
            return null;
        }
    }

    /**
     * Handles an error by capturing the error state, logging details, and returning structured info.
     *
     * @param {object} page - The Playwright page instance.
     * @param {Error} error - The error that occurred.
     * @param {string} context - Additional context for the error.
     * @returns {Promise<object>} Structured error information.
     */
    static async handleError(page, error, context) {
        try {
            // Capture the error state (screenshot, HTML, console logs)
            const state = await this.captureErrorState(page, error, context);

            // Attempt to retrieve the page URL; if unavailable, default to an empty string.
            let pageUrl = '';
            try {
                pageUrl = page.url();
            } catch (urlError) {
                logger.warn('Unable to get page URL:', urlError);
            }
            logger.error('Operation failed', error, {
                context,
                url: pageUrl,
                timestamp: new Date().toISOString()
            });

            return {
                message: error.message,
                context,
                timestamp: new Date().toISOString(),
                url: pageUrl,
                stateCaptured: !!state
            };
        } catch (handleError) {
            logger.error('Error in handleError:', handleError);
            return {
                message: error.message,
                context,
                timestamp: new Date().toISOString(),
                url: '',
                stateCaptured: false
            };
        }
    }

    /**
     * Wraps an asynchronous function with error handling.
     * If the function throws an error, handleError is invoked.
     *
     * @param {object} page - The Playwright page instance.
     * @param {string} context - Additional context for the error.
     * @param {Function} fn - The asynchronous function to execute.
     * @returns {Promise<*>} The result of the function, or structured error info if it fails.
     */
    static async withErrorHandler(page, context, fn) {
        try {
            return await fn();
        } catch (error) {
            return await this.handleError(page, error, context);
        }
    }
}

module.exports = ErrorHandler;
