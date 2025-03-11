// src/utils/retry.js
const logger = require('./logger');

class RetryError extends Error {
    constructor(message, lastError) {
        super(message);
        this.name = 'RetryError';
        this.lastError = lastError;
    }
}

/**
 * Retries an asynchronous operation with configurable options.
 *
 * @param {Function} fn - The asynchronous function to execute.
 * @param {Object} [options={}] - Configuration options.
 * @param {number} [options.retries=3] - Number of retry attempts.
 * @param {number} [options.delay=1000] - Initial delay in milliseconds before retrying.
 * @param {boolean} [options.exponential=true] - Whether to use exponential backoff.
 * @param {number} [options.maxDelay=30000] - Maximum delay in milliseconds between retries.
 * @param {Function} [options.onRetry=null] - Optional callback invoked after each failed attempt.
 *                 It receives the error and the current attempt number.
 * @param {Function} [options.shouldRetry=null] - Optional callback to determine whether to retry.
 *                 It receives the error and the current attempt number and should return a boolean.
 * @param {string} [options.name='operation'] - A name to identify the operation for logging purposes.
 * @returns {Promise<*>} The result of the successful operation.
 * @throws {RetryError} If all retry attempts fail or if shouldRetry returns false.
 */
async function retry(fn, options = {}) {
    const {
        retries = 3,
        delay = 1000,
        exponential = true,
        maxDelay = 30000,
        onRetry = null,
        shouldRetry = null,
        name = 'operation'
    } = options;

    let lastError;
    
    for (let attempt = 1; attempt <= retries; attempt++) {
        try {
            return await fn();
        } catch (error) {
            lastError = error;

            // If a shouldRetry callback is provided and returns false, abort retries.
            if (typeof shouldRetry === 'function' && !shouldRetry(error, attempt)) {
                throw new RetryError(`Aborted ${name} on attempt ${attempt}: ${error.message}`, error);
            }

            // If this is the last attempt, throw a RetryError.
            if (attempt === retries) {
                throw new RetryError(`Failed ${name} after ${retries} attempts: ${error.message}`, lastError);
            }

            // Calculate the delay time (with exponential backoff if enabled), capped by maxDelay.
            const waitTime = exponential ? Math.min(delay * Math.pow(2, attempt - 1), maxDelay) : delay;
            
            logger.warn(
                `Attempt ${attempt}/${retries} failed for ${name}. Retrying in ${waitTime}ms`,
                { error: error.message }
            );

            // Call the onRetry callback if provided.
            if (typeof onRetry === 'function') {
                await onRetry(error, attempt);
            }

            // Wait for the specified delay before trying again.
            await new Promise(resolve => setTimeout(resolve, waitTime));
        }
    }
}

module.exports = {
    retry,
    RetryError
};
