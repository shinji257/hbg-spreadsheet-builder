const flags = process.argv.slice(2);
let debug = false;
let cmdChoice = null;
let cmdRootFolder = null;
let cmdUploadChoice = null;

if (flags.includes('--debug')) {
	flags.splice(flags.indexOf('--debug'), 1);
	debug = true;
}

if (flags.includes('-source')) {
	const argIndex = flags.indexOf('-source');
	cmdChoice = flags[argIndex + 1];
	flags.splice(argIndex, 2);
}

if (flags.includes('-root')) {
	const argIndex = flags.indexOf('-root');
	cmdRootFolder = flags[argIndex + 1];
	flags.splice(argIndex, 2);
}

if (flags.includes('-upload')) {
	const argIndex = flags.indexOf('-upload');
	cmdUploadChoice = flags[argIndex + 1];
	flags.splice(argIndex, 2);
}

function question(question) {
	return new Promise((resolve, reject) => {
		rl.question(question, (answer) => {
			resolve(answer)
		});
	});
}

const fs = require('fs');
const readline = require('readline');
const { google } = require('googleapis');
const xl = require('excel4node');
const moment = require('moment')
const stream = require('stream');

let conf;

if (fs.existsSync('./conf.json')) {
	conf = require('./conf.json');
}

const listNSP = conf.listNSP || null;
const listNSZ = conf.listNSZ || null;
const listOthers = conf.listOthers || null;

const wb = new xl.Workbook();

const SCOPES = ['https://www.googleapis.com/auth/drive'];
const TOKEN_PATH = 'token.json';
let spreadsheetId;
let driveAPI;
let selectedDrive;

if (fs.existsSync('conf.json')) {
	const config = require('./conf.json');
	spreadsheetId = config.spreadsheetId;
}

const rl = readline.createInterface({
	input: process.stdin,
	output: process.stdout,
});

fs.readFile('credentials.json', (err, content) => {
	if (err) return console.log('Error loading client secret file:', err);

	authorize(JSON.parse(content), choice);
});

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
	const {
		client_secret,
		client_id,
		redirect_uris
	} = credentials.installed;
	const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

	fs.readFile(TOKEN_PATH, (err, token) => {
		if (err) return getAccessToken(oAuth2Client, callback);
		oAuth2Client.setCredentials(JSON.parse(token));

		driveAPI = google.drive({
			version: 'v3',
			auth: oAuth2Client
		});

		callback();
	});
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getAccessToken(oAuth2Client, callback) {
	const authUrl = oAuth2Client.generateAuthUrl({
		access_type: 'offline',
		scope: SCOPES,
	});

	console.log('Authorize this app by visiting this url:', authUrl);

	rl.question('Enter the code from that page here: ', (code) => {
		oAuth2Client.getToken(code, (err, token) => {
			if (err) return console.error('Error retrieving access token', err);
			oAuth2Client.setCredentials(token);
			fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
				if (err) return console.error(err);
				console.log('Token stored to', TOKEN_PATH);
			});

			driveAPI = google.drive({
				version: 'v3',
				auth: oAuth2Client
			});
	
			callback();
		});
	});
}

async function choice() {
	const resp = await driveAPI.drives.list({
		fields: 'drives(id, name)'
	}).catch(console.error);

	const result = resp.data.drives;
	let x = 2;

	let chosen = cmdChoice ? cmdChoice === '0' ? 1 : result.findIndex(e => e.id === cmdChoice) + x : null;

	if (!chosen) {
		console.log('1: Your own drive');
		for (const gdrive of result) {
			console.log(`${x++}: ${gdrive.name} (${gdrive.id})`);
		}
	
		chosen = Number(await question('Enter your choice: '));
	} else {
		x += result.length;
	}

	if (chosen === 1) {
		listDriveFiles();
	} else if (chosen <= x && chosen > 1) {
		selectedDrive = `${result[chosen - 2].name} (${result[chosen - 2].id})`;
		listDriveFiles(result[chosen - 2].id);
	} else {
		if (cmdChoice) cmdChoice = null;
		choice();
	}
}

async function listDriveFiles(driveId = null) {
	if (!listNSP && !listNSZ && !listOthers) {
		console.log('Nothing to add to the spreadsheet')
		process.exit();
	}

	const startTime = moment.now();

	const folderOptions = {
		fields: 'nextPageToken, files(id, name)',
		orderBy: 'name'
	};

	let rootfolder = cmdRootFolder;

	if (!cmdRootFolder) rootfolder = await question('Whats the root folder id: ');

	if (driveId) {
		folderOptions.driveId = driveId;
		folderOptions.corpora = 'drive';
		folderOptions.includeItemsFromAllDrives = true;
		folderOptions.supportsAllDrives = true;
	} else {
		folderOptions.corpora = 'user';
	}

	folderOptions.q = `mimeType = \'application/vnd.google-apps.folder\' and trashed = false`;

	folderOptions.q += ` and \'${rootfolder ? rootfolder : driveId}\' in parents`;

	let res_folders = await retrieveAllFiles(folderOptions).catch(console.error);

	const order = ['base', 'dlc', 'updates', 'Custom XCI', 'Custom XCI JP', 'Special Collection', 'XCI Trimmed'];
	const order_nsz = ['base', 'dlc', 'updates'];
		
	let folders = [];
	let folders_nsz = [];

	if (listNSP) {
		folderOptions.q = `mimeType = \'application/vnd.google-apps.folder\' and trashed = false and \'${res_folders[res_folders.map(e => e.name).indexOf('NSP Dumps')].id}\' in parents`;

		const temp = await retrieveAllFiles(folderOptions).catch(console.error);
	
		const res_nsp = res_folders.concat(temp).filter(folder => order.includes(folder.name));
	
		for (const folder of res_nsp) {
			folders[order.indexOf(folder.name)] = folder
		};

		folders = folders.filter(arr => !!arr);
	
		await goThroughFolders(driveId, folders, ['base', 'dlc', 'updates'], {
			base: 'NSP Base',
			dlc: 'NSP DLC',
			updates: 'NSP Updates',
		});
	} else {
		for (const folder of res_folders.filter(folder => order.includes(folder.name))) {
			folders[order.indexOf(folder.name)] = folder
		};

		folders = folders.filter(arr => !!arr);
	}

	if (listNSZ) {
		folderOptions.q = `mimeType = \'application/vnd.google-apps.folder\' and trashed = false and \'${res_folders[res_folders.map(e => e.name).indexOf('NSZ')].id}\' in parents`;
	
		const res_nsz = (await retrieveAllFiles(folderOptions).catch(console.error)).filter(folder => order_nsz.includes(folder.name));
	
		for (const folder of res_nsz) {
			folders_nsz[order_nsz.indexOf(folder.name)] = folder
		};

		folders_nsz = folders_nsz.filter(arr => arr !== null);

		await goThroughFolders(driveId, folders_nsz, ['base', 'dlc', 'updates'], {
			base: 'NSZ Base',
			dlc: 'NSZ DLC',
			updates: 'NSZ Updates',
		});
	}

	if (listOthers) {
		await goThroughFolders(driveId, folders, ['Custom XCI', 'Custom XCI JP', 'XCI Trimmed', 'Special Collection']);
	}

	if (!fs.existsSync('output/')) fs.mkdirSync('output/');

	await wb.write('output/spreadsheet.xlsx');

	console.log('Generation of NSP spreadsheet completed.');
	console.log(`Took: ${moment.utc(moment().diff(startTime)).format('HH:mm:ss.SSS')}`);

	if (driveId) {
		const driveAnswer = await question(`Write to ${selectedDrive}? [y/n]:`);
		if (['y', 'Y', 'yes', 'yeS', 'yEs', 'yES', 'Yes', 'YeS', 'YEs', 'YES'].includes(driveAnswer)) {
			writeToDrive(driveId);
		} else {
			writeToDrive();
		}
	} else {
		writeToDrive();
	}
}

function goThroughFolders(driveId, folders, includeIndex, nameTable = null) {
	return new Promise(async (resolve, reject) => {
		if (!driveId || !folders || !includeIndex) reject('Missing parameters');

		for (const folder of folders) {
			if (!includeIndex.includes(folder.name)) continue;
	
			if (debug) console.log(folder.name);
	
			if (nameTable) {
				const folder_mod = {
					name: nameTable[folder.name],
					id: folder.id
				};

				await addToWorkbook(folder_mod, driveId);
			} else {
				await addToWorkbook(folder, driveId);
			}
		}
		resolve();
	});
}

async function addToWorkbook(folder, driveId = null) {
	return new Promise(async (resolve, reject) => {
		if (!folder) reject('No folder given');

		const options = {
			fields: 'nextPageToken, files(id, name, size, webContentLink, modifiedTime, md5Checksum)',
			orderBy: 'name',
			pageSize: 1000,
			q: `\'${folder.id}\' in parents and trashed = false and not mimeType = \'application/vnd.google-apps.folder\'`
		};

		const sheet = wb.addWorksheet(folder.name);
	
		if (driveId) {
			options.driveId = driveId;
			options.corpora = 'drive';
			options.includeItemsFromAllDrives = true;
			options.supportsAllDrives = true;
		} else {
			options.corpora = 'user';
		}
	
		files = await retrieveAllFiles(options).catch(console.error);
	
		if (files.length) {
			if (debug) console.log(`Files in ${folder.name}:`);

			const columns = [
				{ width: 93, name: 'Name' },
				{ width: 18, name: 'Date updated' },
				{ width: 12, name: 'Size' },
				{ width: 20, name: 'Hash' },
				{ width: 95, name: 'URL' },
			]

			for (let entry in columns) {
				entry = Number(entry);
				sheet.column(entry + 1).setWidth(columns[entry].width);
				sheet.cell(1, entry + 1).string(columns[entry].name);
			}

			sheet.row(1).freeze();
			
			let i = 2;
			for (const file of files) {
				if (debug) console.log(`${file.name} (${file.id})`);
				
				sheet.cell(i,1).string(file.name);
				sheet.cell(i,2).string(moment(file.modifiedTime).format('M/D/YYYY H:m:s'));
				sheet.cell(i,3).string(file.size);
				sheet.cell(i,4).string(file.md5Checksum);
				sheet.cell(i,5).string(file.webContentLink);
				i++;
			};
		} else {
			console.log('No files found.');
		}
		resolve();
	});
}

async function writeToDrive(driveId = null) {
	let answer = cmdUploadChoice;
	
	if (!cmdUploadChoice) answer = await question('Do you want to upload the spreadsheet to your google drive? [y/n]: ');

	if (answer === 'y') {
		await doUpload(driveId)
	}

	process.stdout.write('\nPress any key to exit...');

	process.stdin.setRawMode(true);
	process.stdin.resume();
	process.stdin.on('data', process.exit.bind(process, 0));
}

async function doUpload(driveId = null) {
	return new Promise(async (resolve, reject) => {
		const buf = Buffer.from(fs.readFileSync('output/spreadsheet.xlsx'), 'binary');
		const buffer = Uint8Array.from(buf);
		var bufferStream = new stream.PassThrough();
		bufferStream.end(buffer);
		const media = {
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			body: bufferStream,
		};
	
		if (spreadsheetId) {	
			console.log('Updating the spreadsheet on the drive...');
	
			await driveAPI.files.update({
				fileId: spreadsheetId,
				media
			}).catch(console.error);	  
		} else {
			console.log('Creating the spreadsheet on the drive...')
	
			const fileMetadata = {
				name: '／hbg／ - Donator\'s Spreadsheet 3.0'
			};
	
			if (driveId) {
				fileMetadata.parents = [driveId];
			}
	
			const file = await driveAPI.files.create({
				resource: fileMetadata,
				media,
				fields: 'id'
			}).catch(console.error);
	
			const config = {
				spreadsheetId: file.data.id
			};
	
			fs.writeFileSync('conf.json', JSON.stringify(config, null, '\t'));
		}
	
		console.log('Done!');
		resolve();
	});
}

function retrieveAllFiles(options) {
	return new Promise(async (resolve, reject) => {
		const result = await retrievePageOfFiles(options, []).catch(console.error);
	
		resolve(result);
	});
}

function retrievePageOfFiles(options, result) {
	return new Promise(async (resolve, reject) => {
		const resp = await driveAPI.files.list(options).catch(console.error);
	
		result = result.concat(resp.data.files);
	
		if (resp.data.nextPageToken) {
			options.pageToken = resp.data.nextPageToken;
	
			const res = await retrievePageOfFiles(options, result).catch(console.error);
			resolve(res);
		} else {
			resolve(result);
		}
	});
}