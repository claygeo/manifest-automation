const path = require('path');
const fs = require('fs');
const logger = require('../utils/logger');

class ConfigManager {
    constructor() {
        this.config = null;
        this.load();
    }

    load() {
        try {
            // Load default config
            const defaultConfigPath = path.join(__dirname, 'default.json');
            this.config = require(defaultConfigPath);

            // Check for environment-specific config
            const envConfig = process.env.NODE_ENV || 'development';
            const envConfigPath = path.join(__dirname, `${envConfig}.json`);
            
            if (fs.existsSync(envConfigPath)) {
                const envSettings = require(envConfigPath);
                this.config = this.mergeConfigs(this.config, envSettings);
                logger.info(`Loaded environment config for: ${envConfig}`);
            }

            // Load local config if exists (for user-specific settings)
            const localConfigPath = path.join(__dirname, 'local.json');
            if (fs.existsSync(localConfigPath)) {
                const localSettings = require(localConfigPath);
                this.config = this.mergeConfigs(this.config, localSettings);
                logger.info('Loaded local config settings');
            }

            // Validate configuration
            this.validateConfig();

            logger.info('Configuration loaded successfully');
        } catch (error) {
            logger.error('Error loading configuration:', error);
            throw error;
        }
    }

    mergeConfigs(base, override) {
        return {
            ...base,
            ...override,
            browser: { ...base.browser, ...override.browser },
            retry: { ...base.retry, ...override.retry },
            paths: { ...base.paths, ...override.paths },
            excel: { ...base.excel, ...override.excel },
            validation: { ...base.validation, ...override.validation },
            urls: { ...base.urls, ...override.urls },
            search: { ...base.search, ...override.search }
        };
    }

    validateConfig() {
        const requiredFields = [
            'browser.channel',
            'retry.attempts',
            'paths.screenshots',
            'paths.logs',
            'excel.supportedFormats',
            'validation.requiredFields'
        ];

        for (const field of requiredFields) {
            const value = this.get(field);
            if (value === undefined) {
                throw new Error(`Missing required config field: ${field}`);
            }
        }
    }

    get(path) {
        return path.split('.').reduce((obj, key) => obj && obj[key], this.config);
    }

    set(path, value) {
        const parts = path.split('.');
        const last = parts.pop();
        const obj = parts.reduce((obj, key) => obj[key] = obj[key] || {}, this.config);
        obj[last] = value;
        logger.info(`Updated config value: ${path}`);
    }
}

module.exports = new ConfigManager();