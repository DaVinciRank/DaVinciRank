# DaVinciRank Science Olympiad Tournament Scoring System

## New tournamament Setup

1. In google drive, make a separate folder to be used for tournament scoring (ex `[Tournament Name] [Division] [Tournament Date] Tournament Scoring System`)

2. Make a copy of the master tournamament scoring sheet and place in folder.
  - Latest Version located (2025-26) [here](https://docs.google.com/spreadsheets/d/1SJBLsjXMwXmE-kKM4ijFRyI2IS92urPuC9KIYr3jdEE/copy)

3. Follow instructions on google sheets for rest of setup
- If you get prompt `Authorization Requried`, follow these [instructions](#authorization-required-to-run-script)

4. For build event template files, you can pull from https://www.soinc.org/scoresheets yourself or go to this google drive link and use `Control-C` and `Control-V` to copy all files to your folder at once.
  - Note: Wait for step `4. Create Grading Scoresheets` to do this as it required a specific folder you need to copy those into. Also make sure that the exact event name appears (ie `Robot Tour` not `RobotTour`) in the name of the files, otherwise it will not be copied over to event folders. You may also choose to upload event rules or any other pertinent files here as long it has the event name in the file name.
  - B-Division: https://drive.google.com/drive/folders/1TvcigTrrJmieDnZyEeXOBPzURXXQNhvF
  - C-Division: https://drive.google.com/drive/folders/1uMEusLDGLVHCZrM9XIax6LKqbxAZ9NqU

4. For tournament awards slideshow, make a copy of this presentation and place in same folder as before. Make sure to have exact text `Tournament Medals` in name of file.
  - Latest Version located [here](https://docs.google.com/presentation/d/1LD4XxnMNLWeLIUADrXtVydBi8DByWcr97SnKWymNKxs/copy)

## Usage

[Coming Soon]

### Create Only Event Tabs

### Create Event Spreadsheets

### Create Grading Scoresheets

### Share Scoring Folder with Supervisors

### Create Slides Presentation

## Contributing

### Setup

1. Install clasp

```
npm install -g @google/clasp
```

2. Login to clasp

```
clasp login
```

3. Then enable the Google Apps Script API: https://script.google.com/home/usersettings

4. Find the `scriptId` from the google sheets app script (go to `Project Settings` and copy `Script ID`)

5. Clone the script code to local

```
clasp clone <scriptId>
```

### Pushing Changes to App Script

This pushes any local code to the google scripts project

```
clasp push
```

# Development

## Prerequisites

- Node.js
- google/clasp
  - Global installation is recommended

## Getting Started

### Clone the repository

```
git clone https://github.com/DaVinciRank/DaVinciRank.git
```

### Install dependencies

```
npm install
```

### Development and build project

```
npm run build
```

### Push

```
npm run push
```

## FAQ

### Authorization Required to Run Script
If you get this message, you need to grant permissions to your Google Apps Script.
1. Click on `Review permissions`,
2. Select your Google account (select Google account that you used to create this script, otherwise it won't work as you expected)
3. This will show `Google hasn’t verified this app` page. Click on `Advanced` and click on `Go to DaVinciRank Tournament Scoring System (unsafe)`
4. Finally, click on `Allow` to grant permission to your script.

## Google Apps Script Resources

- https://developers.google.com/apps-script/guides/clasp
- https://github.com/google/clasp/blob/master/docs/typescript.md
- https://developers.google.com/apps-script/guides/support/best-practices
- https://gsuite-developers.googleblog.com/2015/12/advanced-development-process-with-apps.html
- http://googleappsscript.blogspot.com/2010/06/optimizing-spreadsheet-operations.html
- https://freedium.cfd/https://medium.com/geekculture/the-ultimate-guide-to-npm-modules-in-google-apps-script-a84545c3f57c
