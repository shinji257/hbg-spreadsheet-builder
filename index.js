const flags = process.argv.slice(2);
let debug = false;

if (flags.includes('debug') || flags.includes('-debug') || flags.includes('--debug')) debug = true;


const fs = require('fs');
const readline = require('readline');
const { google } = require('googleapis');
const xl = require('excel4node');
const moment = require('moment')
const stream = require('stream');

const wb = new xl.Workbook();

const workbook = [];

const SCOPES = ['https://www.googleapis.com/auth/drive'];
const TOKEN_PATH = 'token.json';
let spreadsheetId;
let driveAPI;

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

		callback(driveAPI);
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
		rl.close();
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
	
			callback(driveAPI);
		});
	});
}

async function choice(driveAPI) {
	const resp = await driveAPI.drives.list({
		fields: 'drives(id, name)'
	}).catch(console.error);

	const result = resp.data.drives;

	console.log('1: Your own drive');
	let x = 2;
	for (const gdrive of result) {
		console.log(`${x}: ${gdrive.name} (${gdrive.id})`);
		x++;
	}

	rl.question('Enter your choice: ', chosen => {
		if (chosen === '1') {
			listDriveFiles(driveAPI);
		} else if (chosen <= x && chosen > 1) {
			listDriveFiles(driveAPI, result[chosen - 2].id);
		} else {
			choice(driveAPI);
		}
	});
}

async function listDriveFiles(driveAPI, driveId = null) {
	const startTime = moment.now();

	const folderOptions = {
		fields: 'nextPageToken, files(id, name)',
		orderBy: 'name',
		q: 'name = \'hbg\''
	};

	if (driveId) {
		folderOptions.driveId = driveId;
		folderOptions.corpora = 'drive';
		folderOptions.includeItemsFromAllDrives = true;
		folderOptions.supportsAllDrives = true;
	} else {
		folderOptions.corpora = 'user';
	}

	let res = await retrieveAllFiles(folderOptions, driveAPI).catch(console.error);

	folderOptions.q = `mimeType = \'application/vnd.google-apps.folder\' and \'${res[0].id}\' in parents`;

	let res_folders = await retrieveAllFiles(folderOptions, driveAPI).catch(console.error);

	const order = ['base', 'dlc', 'updates', 'Custom XCI', 'Custom XCI JP', 'Special Collection', 'XCI Trimmed'];

	folderOptions.q = `mimeType = \'application/vnd.google-apps.folder\' and \'${res_folders[res_folders.map(e => e.name).indexOf('NSP Dumps')].id}\' in parents`;

	const temp = await retrieveAllFiles(folderOptions, driveAPI).catch(console.error);

	res_folders = res_folders.concat(temp);

	let unsorted = res_folders
		.filter(folder => order.includes(folder.name));

	let folders = [];

	for (const folder of unsorted) {
		folders[order.indexOf(folder.name)] = folder
	};

	folders = folders.filter(arr => arr !== null);

	const order_nsz = ['base', 'dlc', 'updates'];

	folderOptions.q = `mimeType = \'application/vnd.google-apps.folder\' and \'${res_folders[res_folders.map(e => e.name).indexOf('NSZ')].id}\' in parents`;

	const res_nsz = await retrieveAllFiles(folderOptions, driveAPI).catch(console.error);

	let unsorted_nsz = res_nsz
		.filter(folder => order_nsz.includes(folder.name));

	let folders_nsz = [];

	for (const folder of unsorted_nsz) {
		folders_nsz[order_nsz.indexOf(folder.name)] = folder
	};

	folders_nsz = folders_nsz.filter(arr => arr !== null);

	let x = 0;
	for (const folder of folders) {
		if (!['base', 'dlc', 'updates'].includes(folder.name)) continue;

		if (debug) console.log(folder.name);

		const table = {
			base: 'NSP Base',
			dlc: 'NSP DLC',
			updates: 'NSP Updates',
		};

		const folder_mod = {
			name: table[folder.name],
			id: folder.id
		};

		await addToWorkbook(folder_mod, driveAPI, driveId);

		x++;
	}

	for (const folder of folders_nsz) {
		if (!['base', 'dlc', 'updates'].includes(folder.name)) continue;

		if (debug) console.log(folder.name);

		const table = {
			base: 'NSZ Base',
			dlc: 'NSZ DLC',
			updates: 'NSZ Updates',
		};

		const folder_mod = {
			name: table[folder.name],
			id: folder.id
		};

		await addToWorkbook(folder_mod, driveAPI, driveId);

		x++;
	}

	for (const folder of folders) {
		if (!['Custom XCI', 'Custom XCI JP', 'XCI Trimmed', 'Special Collection'].includes(folder.name)) continue;

		if (debug) console.log(folder.name);

		await addToWorkbook(folder, driveAPI, driveId);

		x++;
	}

	if (!fs.existsSync('output/')) fs.mkdirSync('output/');

	wb.write('output/spreadsheet.xlsx');

	console.log('Generation of NSP spreadsheet completed.');
	console.log(`Took: ${moment.utc(moment().diff(startTime)).format("HH:mm:ss.SSS")}`);

	writeToDrive(driveAPI);
}

async function addToWorkbook(folder, driveAPI, driveId = null) {
	return new Promise(async (resolve, reject) => {
		const options = {
			fields: 'nextPageToken, files(id, name, size, webContentLink, modifiedTime, md5Checksum)',
			orderBy: 'name',
			pageSize: 1000,
			q: `\'${folder.id}\' in parents and not mimeType = \'application/vnd.google-apps.folder\'`
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
	
		files = await retrieveAllFiles(options, driveAPI).catch(console.error);
	
		if (files.length) {
			if (debug) console.log(`Files in ${folder.name}:`);
	
			sheet.column(1).setWidth(93);
			sheet.column(2).setWidth(18);
			sheet.column(3).setWidth(12);
			sheet.column(4).setWidth(20);
			sheet.column(5).setWidth(95);
	
			sheet.cell(1,1).string('Name');
			sheet.cell(1,2).string('Date updated');
			sheet.cell(1,4).string('Size');
			sheet.cell(1,4).string('Hash');
			sheet.cell(1,5).string('URL');
			
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

function writeToDrive(driveAPI, driveId = null) {
	rl.question('Do you want to upload the spreadsheet to your google drive? [y/n]: ', async (answer) => {
		if (answer === 'y') { 
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
		}
		process.stdout.write('\nPress any key to exit...');

		process.stdin.setRawMode(true);
		process.stdin.resume();
		process.stdin.on('data', process.exit.bind(process, 0));
	});
}

function retrieveAllFiles(options, driveAPI) {
	return new Promise(async (resolve, reject) => {
		const result = await retrievePageOfFiles(options, [], driveAPI).catch(console.error);
	
		resolve(result);
	});
}

function retrievePageOfFiles(options, result, driveAPI) {
	return new Promise(async (resolve, reject) => {
		const resp = await driveAPI.files.list(options).catch(console.error);
	
		result = result.concat(resp.data.files);
	
		if (resp.data.nextPageToken) {
			options.pageToken = resp.data.nextPageToken;
	
			const res = await retrievePageOfFiles(options, result, driveAPI).catch(console.error);
			resolve(res);
		} else {
			resolve(result);
		}
	});
}