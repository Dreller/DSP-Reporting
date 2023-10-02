<h1 align="center">
DSP Reporting
</h1>
<h4 align="center">A light and simple utility for reporting using SharePoint Lists.</h4>

<div style="text-align:center;width:100%;">
<img src="https://img.shields.io/badge/Microsoft_SharePoint-0078D4?style=for-the-badge&logo=microsoft-sharepoint&logoColor=white">
<img src="https://img.shields.io/github/issues/dreller/DSP-Reporting.svg">
<img src="https://img.shields.io/github/license/dreller/DSP-Reporting.svg"> 
</div>

<p align="center">
  <a href="#key-features">Key Features</a> •
  <a href="#installation">Installation</a> •
  <a href="#how-to-use">How To Use</a> •
  <a href="#download">Download</a> •
  <a href="#license">License</a>
</p>


## Key Features

- Build reports using the 3S: Select, Sort and Show.
- Share a single URL to view the report.

## Installation

**Instructions to set the Report List**

1. Download the latest release from the [Releases](https://github.com/Dreller/DSP-Reporting/releases) page.
1. Extract the content on your computer.
1. Copy the file `bin/dconf.template.js`, rename to `bin/dconf.js`.  See the [Configuration file](configuration-file) to know how to set it up.
1. In your SharePoint Site, upload all files in a new Document Library or in a new Folder in an existing Library.
1. 

### Configuration file

The Configuration file is named `dconf.js`, under the `bin` folder.
Open the file with your text editor and replace the values following these instructions.

**dconf.js**
```
/**
 * Configuration File
 */

// SharePoint Site URL
var _URL = "https://123.sharepoint.com/sites/Mysite";

// List for saving Reports
var _REPORTLIST = "Reports";
```

`_URL`: URL to the root of the site your are installing DSP Reporting.
`_REPORTLIST`: Name of the SharePoint List where Reports will be saved.


## How To Use

...


## Download

Download the latest release from the [Releases](https://github.com/Dreller/DSP-Reporting/releases) page.



## License

...


---