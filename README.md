# Science Olympiad Tournament Scoring System

## Usage

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

# Deveopment

## Prerequisites

- Node.js
- google/clasp
  - Global installation is recommended

## Getting Started

### Clone the repository

```
git clone https://github.com/rohanskariah/DaVinciRank
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

## Google Apps Script Resources

- https://developers.google.com/apps-script/guides/clasp
- https://github.com/google/clasp/blob/master/docs/typescript.md
- https://developers.google.com/apps-script/guides/support/best-practices
- https://gsuite-developers.googleblog.com/2015/12/advanced-development-process-with-apps.html
- http://googleappsscript.blogspot.com/2010/06/optimizing-spreadsheet-operations.html
- https://freedium.cfd/https://medium.com/geekculture/the-ultimate-guide-to-npm-modules-in-google-apps-script-a84545c3f57c