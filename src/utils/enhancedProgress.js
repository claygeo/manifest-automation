const logger = require('./logger');
const cliProgress = require('cli-progress');
const colors = require('ansi-colors');

class EnhancedProgressTracker {
    constructor(totalSteps, name) {
        this.totalSteps = totalSteps;
        this.name = name;
        this.startTime = Date.now();
        this.steps = [];
        this.currentStep = 0;
        this.errors = [];
        this.warnings = [];

        // Create a multi-bar container
        this.multibar = new cliProgress.MultiBar({
            clearOnComplete: false,
            hideCursor: true,
            format: this.createProgressBarFormat(),
        }, cliProgress.Presets.shades_classic);

        // Main progress bar for overall progress
        this.mainBar = this.multibar.create(totalSteps, 0, {
            name: name,
            status: 'Initializing...'
        });

        // Step-specific progress bar
        this.stepBar = this.multibar.create(100, 0, {
            name: 'Current Step',
            status: 'Waiting to start...'
        });
    }

    createProgressBarFormat() {
        return colors.cyan('{bar}') + ' | {percentage}% | {name} | {status} | {duration_formatted}';
    }

    formatDuration(ms) {
        const seconds = Math.floor(ms / 1000);
        const minutes = Math.floor(seconds / 60);
        const hours = Math.floor(minutes / 60);
        return `${hours}h:${minutes % 60}m:${seconds % 60}s`;
    }

    /**
     * Returns the currently active step object.
     */
    getCurrentStep() {
        if (this.steps.length > 0) {
            return this.steps[this.steps.length - 1];
        }
        return null;
    }

    updateMainProgress(status) {
        const overallDuration = Date.now() - this.startTime;
        this.mainBar.update(this.currentStep, {
            status,
            duration_formatted: this.formatDuration(overallDuration)
        });
    }

    updateStepProgress(percentage, status) {
        // If there is a current step, compute its duration; otherwise fallback to overall startTime.
        const currentStep = this.getCurrentStep();
        const stepStart = currentStep ? currentStep.startTime : this.startTime;
        const stepDuration = Date.now() - stepStart;
        this.stepBar.update(percentage, {
            status,
            duration_formatted: this.formatDuration(stepDuration)
        });
    }

    addStep(stepName) {
        const step = {
            name: stepName,
            startTime: Date.now(),
            status: 'In Progress'
        };
        this.steps.push(step);

        // Update progress bars
        this.updateMainProgress(`Processing: ${stepName}`);
        this.updateStepProgress(0, 'Starting...');

        logger.info(`→ ${stepName} started`);
    }

    completeStep(stepName) {
        const step = this.steps.find(s => s.name === stepName);
        if (step) {
            step.status = 'Completed';
            step.endTime = Date.now();
            step.duration = step.endTime - step.startTime;
        }

        this.currentStep++;
        this.updateMainProgress(`Completed: ${stepName}`);
        this.updateStepProgress(100, 'Done');

        logger.info(`✓ ${stepName} completed successfully`);
    }

    addError(stepName, error) {
        this.errors.push({
            step: stepName,
            error: error.message,
            timestamp: new Date().toISOString()
        });
        this.updateStepProgress(100, colors.red('Failed'));
        logger.error(`✘ ${stepName} failed:`, error);
    }

    addWarning(stepName, message) {
        this.warnings.push({
            step: stepName,
            message,
            timestamp: new Date().toISOString()
        });
        logger.warn(`⚠ ${stepName}: ${message}`);
    }

    finish() {
        this.multibar.stop();
        this.logSummary();
    }

    logSummary() {
        const overallDuration = Date.now() - this.startTime;
        console.log('\n' + colors.bold('=== Process Summary ==='));
        console.log(colors.cyan(`Process: ${this.name}`));
        console.log(colors.cyan(`Total Duration: ${this.formatDuration(overallDuration)}`));
        console.log(colors.cyan(`Steps Completed: ${this.currentStep}/${this.totalSteps}`));
        console.log('');
        this.steps.forEach(step => {
            const durationStr = step.duration ? this.formatDuration(step.duration) : 'In Progress';
            console.log(colors.green(`• ${step.name}: ${step.status} | Duration: ${durationStr}`));
        });
        
        if (this.errors.length > 0) {
            console.log(colors.red('\nErrors:'));
            this.errors.forEach(error => {
                console.log(colors.red(`  • ${error.step}: ${error.error}`));
            });
        }

        if (this.warnings.length > 0) {
            console.log(colors.yellow('\nWarnings:'));
            this.warnings.forEach(warning => {
                console.log(colors.yellow(`  • ${warning.step}: ${warning.message}`));
            });
        }
    }
}

module.exports = EnhancedProgressTracker;
