# Manifest Simple

[![GitHub license](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/kmaclip/manifest-simple/blob/main/LICENSE)
[![GitHub issues](https://img.shields.io/github/issues/kmaclip/manifest-simple)](https://github.com/kmaclip/manifest-simple/issues)

**Manifest Simple** is a desktop application built with [Electron](https://www.electronjs.org/) designed to simplify the creation, editing, or management of manifest files (e.g., for web apps, shipping, or software packaging). [Insert a one-sentence description of what your tool specifically does here.]

## Features

- **User-Friendly Interface**: Intuitive GUI for managing manifest files.
- **Cross-Platform**: Runs on Windows, macOS, and Linux via Electron.
- **Customizable**: [Add specific features, e.g., "Supports JSON validation" or "Generates manifests from templates"].
- **Lightweight**: Minimal dependencies for easy setup.

## Prerequisites

Before you begin, ensure you have the following installed:
- [Node.js](https://nodejs.org/) (v16.x or later recommended)
- [npm](https://www.npmjs.com/) (comes with Node.js)
- [Git](https://git-scm.com/) (to clone the repository)

## Installation

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/kmaclip/manifest-simple.git
   cd manifest-simple
   
2. Install Dependencies: Run the following command to install required Node.js packages:
	npm install
	
3. Run the Application: Start the app in development mode:
	npm start
	
## Usage

1. Launch the App: 
	Run npm start to open the application.
	An Electron dashboard will popup with the filled  out data that can be edited.
	The Electron dashboard pulls the data from the excel sheet in the root folder.
	
2. Electron Dashboard
	It is vital that all the information is filled out. All of the necessary information that the script needs to create a manifest for Sweed is obtain from the Electron dashboard.
	After pressing 'Approve' the manifest will be automatically created, and also downloaded. 
	

	
## Project Structure

manifest-simple/
├── node_modules/         # Dependencies (ignored by Git)
├── src/                 # Source code (adjust based on your structure)
│   ├── main.js          # Electron main process
│   └── renderer.js      # Renderer process (UI logic)
├── package.json         # Project metadata and scripts
├── .gitignore           # Files ignored by Git
└── README.md            # This file

## Troubleshooting

	App doesn’t start: Ensure Node.js and npm are installed, and run npm install again.
	Large file errors: The repository excludes node_modules—do not commit it. Use .gitignore.
	For other issues, check the Issues page or file a new one.
	
