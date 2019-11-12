# HBG spreadhseet generator

This is a small node.js script that will allow you to easily generate a spreadsheet of your own.


# Installation

## Requirements
- [NodeJS](https://nodejs.org/en/) (v12.X preferably even newer)
- [YarnPKG](https://yarnpkg.com/lang/en/)
- A credentials file of the Google account with access to your stash (Get one [here](https://developers.google.com/drive/api/v3/quickstart/nodejs))
- A few brain cells

## How to install

1. Open a command prompt
2. Navigate to the folder containing the generator
3. Run `yarn install`

# Usage

Here I will explain how to use this tool in a bit more detail.

## Interactive

Just run `node index.js` and the rest will be gone through interactively.

## Automated

Here I will explain on how to automate this tool for usage with cron etc.

### Command-line flags

|Flag|Required|Argument value|
|--|--|--|
|`-auto`|❌||
|`-debug`|❌||
|`-source`|❌|Number or Google Shared Drive ID|
|`-root`|❌|Google Folder ID|
|`-upload`|❌|Either `y` or `n`|
|`-uploadDrive`|❌|Either `y` or `n`|

#### Auth flag note
The auth flag will set every file the script goes through to be accessible by anyone with the link.
This essentially enables us to share the files inside a shared drive as if it were a shared folder.
If you use your own `My Drive` then please share the folder and use Tinfoil.io for setting up the location.

### Examples of automation
Now that we established the flags lets see some examples:

`node index.js -auto -source 1`

This will look for folders called `NSP Dump`, `NSZ`, `XCI Trimmed`, `Custom XCI` in the provided shared drive's root directory. (You can configure what types of files it will look for using the `config.json`)

`node index.js -auto -source SHARED_DRIVE_ID`

This will look for the same folders in the shared drives root directory if not specified otherwise

`node index.js -auto -source SHARED_DRIVE_ID -root FOLDER_ID`

This will do almost the same as not providing a folder id. The only difference being that it will look for the folders in the provided folder instead of the shared drive's root directory

`node index.js -auto -source SHARED_DRIVE_ID -upload y`

This will upload the resulting index to your Google `mydrive` location

`node index.js -auto -source SHARED_DRIVE_ID -uploadDrive y -upload y`

This will upload the resulting file to the source shared drive if one was provided. This cant be used without also setting upload to `y`
