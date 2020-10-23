const progArgs = process.argv.slice(2);
const flags = {};
flags.auto = getArgument('auto', true, false);
flags.debug = getArgument('debug', true, false);
flags.choice = getArgument('source', false);
flags.root = getArgument('root', false);
flags.upload = getArgument('upload', false);
flags.uploadDrive = getArgument('uploadDrive', false);

function getArgument(name, isFlag, defaultValue = null) {
	if (progArgs.includes(`-${name}`)) {
		const index = progArgs.indexOf(`-${name}`);
		if (!isFlag) {
			var argValue = progArgs[index + 1];
		}
		progArgs.splice(index, isFlag ? 1 : 2);
		return isFlag ? true : argValue;
	}
	return defaultValue;
}

function question(question) {
	return new Promise((resolve, reject) => {
		rl.question(question, (answer) => {
			resolve(answer);
		});
	});
}

const fs = require('fs');
const readline = require('readline');
const { google } = require('googleapis');
const xl = require('excel4node');
const moment = require('moment');
const path = require('path');
const cliProgress = require('cli-progress');
const { Worker } = require('worker_threads');

let conf = {};

if (fs.existsSync('./conf.json')) {
	conf = require('./conf.json');
}

conf.listNSP = conf.listNSP || false;
conf.listNSZ = conf.listNSZ || false;
conf.listXCI = conf.listXCI || false;
conf.listCustomXCI = conf.listCustomXCI || false;
conf.spreadsheetId = conf.spreadsheetId || '';

const progBar = new cliProgress.SingleBar({
	format: 'Adding files: [{bar}] {percentage}% | ETA: {eta}s | {value}/{total} files',
	etaBuffer: 100
}, cliProgress.Presets.shades_classic);

const folderBar = new cliProgress.SingleBar({
	format: 'Getting folders: [{bar}] {percentage}% | ETA: {eta}s | {value}/{total} folders',
	etaBuffer: 100
}, cliProgress.Presets.shades_classic);

setInterval(() => {
	if (progBar.isActive) progBar.updateETA();
	if (folderBar.isActive) folderBar.updateETA();
}, 1000);

const wb = new xl.Workbook();

const SCOPES = ['https://www.googleapis.com/auth/drive'];
const TOKEN_PATH = 'gdrive.token';
let driveAPI;
let selectedDrive;

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
	if (credentials.type && credentials.type === "service_account") {
		const {
			client_email,
			private_key
		} = credentials;

		const jwtClient = new google.auth.JWT(
			client_email,
			null,
			private_key,
			['https://www.googleapis.com/auth/drive']);
	
		fs.readFile(TOKEN_PATH, (err, token) => {
			if (err) return getAccessTokenJWT(jwtClient, callback);
			jwtClient.setCredentials(JSON.parse(token));
	
			driveAPI = google.drive({
				version: 'v3',
				auth: jwtClient
			});
	
			callback();
		});
	} else {
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

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.JWT} jwtClient The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getAccessTokenJWT(jwtClient, callback) {
	jwtClient.authorize(function (err, tokens) {
		if (err) return console.error(err);
		
		jwtClient.setCredentials(tokens);
		fs.writeFile(TOKEN_PATH, JSON.stringify(tokens), (err) => {
			if (err) return console.error(err);
			console.log('Token stored to', TOKEN_PATH);
		});

		driveAPI = google.drive({
			version: 'v3',
			auth: jwtClient
		});

		callback();
	});
}

async function choice() {
	const drives = await retrieveAllDrives({
		fields: 'nextPageToken, drives(id, name)'
	}).catch(console.error);
	let x = 1;

	let chosen = flags.choice || null;
	const chosenIsNaN = isNaN(Number(chosen));

	if (chosenIsNaN && chosen !== null) {
		const foundIndex = drives.findIndex(e => e.id === chosen);

		if (foundIndex < 0) chosen = null;
		else chosen = foundIndex + 2;
	}

	chosen = Number(chosen);

	if (!chosen && !flags.auto) {
		console.log('1: Your own drive');
		for (const gdrive of drives) {
			console.log(`${++x}: ${gdrive.name} (${gdrive.id})`);
		}
	
		chosen = Number(await question('Enter your choice: ').catch(console.error));
	} else if (!chosen && flags.auto) {
		console.error('Source argument invalid. Aborting auto.');
		process.exit(1);
	} else {
		x += drives.length;
	}

	if (chosen === 1) {
		listDriveFiles();
	} else if (chosen <= x && chosen > 1) {
		selectedDrive = `${drives[chosen - 2].name} (${drives[chosen - 2].id})`;
		listDriveFiles(drives[chosen - 2].id);
	} else {
		if (flags.choice) flags.choice = null;
		choice();
	}
}

async function listDriveFiles(driveId = null) {
	if (!conf.listNSP && !conf.listNSZ && !conf.listXCI && !conf.listCustomXCI) {
		console.log('Nothing to add to the spreadsheet')
		process.exit();
	}

	const startTime = moment.now();

	const folderOptions = {
		pageSize: 100,
		fields: 'nextPageToken, files(id, name)',
		orderBy: 'name'
	};

	let rootfolder = flags.root;

	if (!rootfolder && !flags.auto) rootfolder = await question('Whats the root folder id: ');
	if (!rootfolder && flags.auto) {
		debugMessage('Invalid root argument. Assuming shared drive as root.');
	}

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

	let res_folders = await retrieveAllFolders(folderOptions).catch(console.error);

	const order = ['base', 'dlc', 'updates', 'Custom XCI', 'Custom XCI JP', 'Special Collection', 'XCI Trimmed'];
	const order_nsz = ['base', 'dlc', 'updates'];
		
	let folders = [];
	let folders_nsz = [];

	if (conf.listNSZ) {
		const nspFolder = res_folders[res_folders.map(e => e.name).indexOf('NSZ')];

		if(nspFolder) {
			folderOptions.q = `mimeType = \'application/vnd.google-apps.folder\' and trashed = false and \'${nspFolder.id}\' in parents`;
		
			const res_nsz = (await retrieveAllFolders(folderOptions).catch(console.error)).filter(folder => order_nsz.includes(folder.name));
		
			for (const folder of res_nsz) {
				folders_nsz[order_nsz.indexOf(folder.name)] = folder
			};
	
			folders_nsz = folders_nsz.filter(arr => arr !== null);
	
			await goThroughFolders(driveId, folders_nsz, ['base', 'dlc', 'updates'], {
				base: 'NSZ Base',
				dlc: 'NSZ DLC',
				updates: 'NSZ Updates',
			});
		} else {
			console.error('No NSZ Folder found');
		}
	}

	if (conf.listNSP) {
		const nszFolder = res_folders[res_folders.map(e => e.name).indexOf('NSP Dumps')];

		if (nszFolder) {
			folderOptions.q = `mimeType = \'application/vnd.google-apps.folder\' and trashed = false and \'${nszFolder.id}\' in parents`;
	
			const temp = await retrieveAllFolders(folderOptions).catch(console.error);
		
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
			console.error('No NSP Folder found');
		}
	} else {
		for (const folder of res_folders.filter(folder => order.includes(folder.name))) {
			folders[order.indexOf(folder.name)] = folder
		};

		folders = folders.filter(arr => !!arr);
	}

	if (conf.listXCI) {
		await goThroughFolders(driveId, folders, ['XCI Trimmed']);
	}

	if (conf.listCustomXCI) {
		const customXCIFolder = folders[folders.map(e => e.name).indexOf('Custom XCI')];

		if (customXCIFolder) {
			folderOptions.q = `mimeType = \'application/vnd.google-apps.folder\' and trashed = false and \'${customXCIFolder.id}\' in parents`;

			const temp = await retrieveAllFolders(folderOptions).catch(console.error);
		
			const res_xci = folders.concat(temp).filter(folder => order.includes(folder.name));
		
			for (const folder of res_xci) {
				folders[order.indexOf(folder.name)] = folder
			};

			folders = folders.filter(arr => !!arr);
		
			await goThroughFolders(driveId, folders, ['Custom XCI', 'Custom XCI JP', 'Special Collection']);
		} else {
			console.error('No Custom XCI folder found');
		}
	}

	if (!fs.existsSync('output/')) fs.mkdirSync('output/');

	wb.write('./output/spreadsheet.xlsx', async (err, stats) => {
		if (err) return console.error(err);

		console.log('Generation of NSP spreadsheet completed.');
		console.log(`Took: ${moment.utc(moment().diff(startTime)).format('HH:mm:ss.SSS')}`);
	
		if (driveId) {
			let driveAnswer = flags.uploadDrive;
	
			if (!driveAnswer && !flags.auto) driveAnswer = await question(`Write to ${rootfolder ? rootfolder : selectedDrive}? [y/n]:`);
			if (!driveAnswer && flags.auto) {
				debugMessage('Invalid uploadDrive argument. Assuming no upload to shared drive.');
				writeToDrive();
			}
			if (['y', 'Y', 'yes', 'yeS', 'yEs', 'yES', 'Yes', 'YeS', 'YEs', 'YES'].includes(driveAnswer)) {
				writeToDrive(driveId);
			} else {
				writeToDrive();
			}
		} else {
			writeToDrive();
		}
	});
}

function goThroughFolders(driveId, folders, includeIndex, nameTable = null) {
	return new Promise(async (resolve, reject) => {
		if (!folders || !includeIndex) reject('Missing parameter');

		for (const folder of folders) {
			if (!includeIndex.includes(folder.name)) continue;
	
			debugMessage(folder.name);
	
			if (nameTable) {
				const folder_mod = {
					name: nameTable[folder.name],
					id: folder.id,
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
			pageSize: 100,
			q: `\'${folder.id}\' in parents and trashed = false and not mimeType = \'application/vnd.google-apps.folder\'`,
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
	
		files = await retrieveAll([folder.id], options, false).catch(console.error);
	
		if (files.length) {
			debugMessage(`Files in ${folder.name}:`);

			const columns = [
				{ width: 93, name: 'Name' },
				{ width: 20, name: 'Date updated' },
				{ width: 15, name: 'Size' },
				{ width: 38, name: 'Hash' },
				{ width: 15, name: 'URL' },
			]

			for (let entry in columns) {
				entry = Number(entry);
				sheet.column(entry + 1).setWidth(columns[entry].width);
				sheet.cell(1, entry + 1).string(columns[entry].name);
			}

			sheet.row(1).freeze();
			
			let i = 2;
			for (const file of files) {
				debugMessage(`${file.name} (${file.id})`);

				const extension = path.extname(file.name);
				if (!['.nsp', '.nsz', '.xci'].includes(extension)) continue;
				
				sheet.cell(i,1).string(file.name);
				sheet.cell(i,2).string(moment(file.modifiedTime).format('M/D/YYYY H:mm:ss'));
				sheet.cell(i,3).string(getFormattedSize(file.size));
				sheet.cell(i,3).comment(`${file.size} B`);
				sheet.cell(i,4).string(file.md5Checksum);
				sheet.cell(i,5).link(file.webContentLink, 'DOWNLOAD');
				i++;
			}
		} else {
			console.log('No files found.');
		}
		resolve();
	});
}

async function writeToDrive(driveId = null) {
	let answer = flags.upload;
	
	if (!answer && !flags.auto) answer = await question('Do you want to upload the spreadsheet to your google drive? [y/n]: ');
	if (!answer && flags.auto) {
		debugMessage('Invalid upload argument. Assuming to not upload the file.');
	}

	if (answer === 'y') {
		await doUpload(driveId);
	}

	if (!flags.auto) {
		process.stdout.write('\nPress any key to exit...');
	
		process.stdin.setRawMode(true);
		process.stdin.resume();
		process.stdin.on('data', process.exit.bind(process, 0));
	} else {
		process.exit(0);
	}
}

async function doUpload(driveId = null) {
	return new Promise(async (resolve, reject) => {
		const media = {
			mimeType: 'application/vnd.ms-excel',
			body: fs.createReadStream('./output/spreadsheet.xlsx'),
		};

		const fileMetadata = {
			mimeType: 'application/vnd.google-apps.spreadsheet',
		}
	
		const requestData = {
			media,
		};

		if (driveId) {
			requestData.driveId = driveId;
			requestData.corpora = 'drive';
			requestData.includeItemsFromAllDrives = true;
			requestData.supportsAllDrives = true;
		}

		if (conf.spreadsheetId) {	
			console.log('Updating the spreadsheet on the drive...');

			requestData.resource = fileMetadata;
			requestData.fileId = conf.spreadsheetId;
	
			await driveAPI.files.update(requestData).catch(console.error);	  
		} else {
			console.log('Creating the spreadsheet on the drive...');
	
			fileMetadata.name = '／hbg／ - Donator\'s Spreadsheet 3.0';
	
			if (driveId) {
				if (flags.root) {
					fileMetadata.parents = [flags.root];
				} else {
					fileMetadata.parents = [driveId];
				}
			}
	
			requestData.resource = fileMetadata;
			requestData.fields = 'id';

			const file = await driveAPI.files.create(requestData).catch(console.error);
	
			conf.spreadsheetId = file.data.id;
	
			fs.writeFileSync('conf.json', JSON.stringify(conf, null, '\t'));
		}
	
		console.log('Done!');
		resolve();
	});
}

function retrieveAll(folderIds, options, recurse = true) {
	return new Promise(async (resolve, reject) => {
		const result = [];

		if (recurse) {
			if (folderIds.length > 0) {
				for (folderId of folderIds) {
					options.q = `\'${folderId}\' in parents and trashed = false and mimeType = \'application/vnd.google-apps.folder\'`;
					result.push(...await retrieveAllFolders(options).catch(reject));
		
					result.push({id: folderId});
				}
			} else {
				options.q = `trashed = false and mimeType = \'application/vnd.google-apps.folder\'`;
				result.push(...await retrieveAllFolders(options).catch(reject));
			}
		} else {
			result.push(folderIds.length > 1 ? folderids : {id: folderIds[0]});
		}

		let promises = [];
		
		//folderBar.start(result.length, 0);

		for (const folder of result) {
			debugMessage(`Getting files from ${folder.id}`);
			promises.push(runFolderWorker({
				folder,
				options
			}));
		}

		const resp = await Promise.all(promises).catch(console.error);

		//folderBar.stop();

		resolve([].concat.apply([], resp.filter(val => val.length > 0)));
	});
}

function retrieveAllFolders(options, result = []) {
	return new Promise(async (resolve, reject) => {
		const resp = await driveAPI.files.list(options).catch(reject);
	
		result = result.concat(resp.data.files);
	
		if (resp.data.nextPageToken) {
			options.pageToken = resp.data.nextPageToken;
	
			const res = await retrieveAllFolders(options, result).catch(reject);
			resolve(res);
		} else {
			resultMap = result.map(v => v.id);
			result = result.filter((v,i) => resultMap.indexOf(v.id) === i);

			let response = [];
			for (const folder of result) {
				options.q = `\'${folder.id}\' in parents and trashed = false and mimeType = \'application/vnd.google-apps.folder\'`;
				delete options.pageToken;
				const resp = await retrieveAllFolders(options).catch(reject);
				response = response.concat(resp);
			}

			response = response.concat(result);

			responseMap = response.map(v => v = v.id);
			response = response.filter((v,i) => responseMap.indexOf(v.id) === i);

			resolve(response);
		}
	});
}

function retrieveAllDrives(options, result = []) {
	return new Promise(async (resolve, reject) => {
		const resp = await driveAPI.drives.list(options).catch(reject);
	
		result = result.concat(resp.data.drives);

		if (resp.data.nextPageToken) {
			options.pageToken = resp.data.nextPageToken;
	
			const res = await retrieveAllDrives(options, result).catch(reject);
			resolve(res);
		} else {
			resolve(result);
		}
	});
}

const sizeSuffix = [
	'B',
	'KB',
	'MB',
	'GB',
	'TB',
	'PB',
	'EB',
	'ZB',
	'YB'
];

function getFormattedSize(size, decimals = 2, round = 0) {
	const tempSize = size / 2**10;
	
	if (tempSize < 1) {
		return `${floorToDecimal(size, decimals)} ${sizeSuffix[round]}`;
	}

	return getFormattedSize(tempSize, decimals, ++round);
}

function floorToDecimal(number, decimals) {
	return Math.floor(number * ( 10 ** decimals )) / 10 ** decimals;
}

function runFolderWorker(workerData) {
	return new Promise((resolve, reject) => {
		const worker = new Worker('./worker.js', { workerData });
		worker.on('message', data => {
			folderBar.increment();
			resolve(data);
		});
		worker.on('error', (err) => {
			console.error(err);
			reject(err)
		});
		worker.on('exit', (code) => {
		if (code !== 0)
			reject(new Error(`Worker stopped with exit code ${code}`));
		})
	});
}

function debugMessage(text) {
	if (flags.debug) {
		console.log(text);
	}
}
