var uuid = require('node-uuid');
var fs = require('fs');
var xmlBuilder = require('xmlBuilder');
var globby = require('globby');
var path = require('path');
var mkdirp = require('mkdirp');
var httpreq = require('httpreq');
var unzip = require('unzip');
var spawnSync = require('child_process').spawnSync;

function pick(arg, defaultValue) {
	return typeof arg == 'undefined' ? defaultValue : arg;
}

function isGuid(str) {
	return /\b[a-fA-F0-9]{8}(?:-[a-fA-F0-9]{4}){3}-[a-fA-F0-9]{12}\b/.test(str);
}

function getGuid(version, guids, guidsPath) {
	if (!guids.hasOwnProperty(version)) {
		guids[version] = uuid.v4().toUpperCase();

		if (guidsPath !== undefined) {
			fs.writeFileSync(guidsPath, JSON.stringify(guids, null, '    '))
		}
	}

	return guids[version];
}

var SimpleMsi = function (options) {
	this.name = pick(options.name, '');
	this.manufacturer = pick(options.manufacturer, '');
	this.arch = pick(options.arch, 'x86');
	this.version = pick(options.version, '0.1.0');

	if (typeof options.guids == 'string') {
		this.guidsPath = options.guids;

		if (fs.existsSync(options.guids)) {
			this.guids = JSON.parse(fs.readFileSync(options.guids, 'utf8'));
		} else {
			this.guids = {};
		}
	} else {
		this.guids = options.guids;
	}

	this.versionGuid = getGuid(this.version, this.guids, this.guidsPath);
	this.guid = getGuid('upgrade', this.guids, this.guidsPath);

	this.programFiles = new ProgramFilesFolder(this, null, this.arch == 'x86' ? 'ProgramFilesFolder' : 'ProgramFiles64Folder');
	this.startMenu = new StartMenuFolder(this, null, 'StartMenu');
	this.desktop = new Desktop(this);
	this.nextId = 0;
};

SimpleMsi.prototype.toWxs = function () {
	var wix = xmlBuilder.create('Wix');
	wix.dec('1.0', 'UTF-8');
	wix.att('xmlns', 'http://schemas.microsoft.com/wix/2006/wi');

	var product = wix.ele('Product', {
		'Id': '*',
		'UpgradeCode': this.guid,
		'Name': this.name,
		'Version': this.version,
		'Manufacturer': this.manufacturer, 
		'Language': '1033'
	});

	var package = product.ele('Package', {
		'InstallerVersion': '405',
		'Compressed': 'yes',
		'Comments': 'Windows Installer Package',
		'Platform': this.arch
	});

	var media = product.ele('Media', {
		'Id': '1',
		'Cabinet': 'product.cab',
		'EmbedCab': 'yes'
	});

	var upgrade = product.ele('Upgrade', {
		'Id': this.guid
	});

	upgrade.ele('UpgradeVersion', {
		'Minimum': this.version,
		'OnlyDetect': 'yes',
		'Property': 'NEWERVERSIONDETECTED'
	});

	upgrade.ele('UpgradeVersion', {
		'Minimum': '0.0.0',
		'Maximum': this.version,
		'IncludeMinimum': 'yes',
		'IncludeMaximum': 'no',
		'Property': 'OLDERVERSIONBEINGUPGRADED'
	});

	var cond = product.ele('Condition', {
		'Message': 'A newer version of this software is already installed.'
	}, 'NOT NEWERVERSIONDETECTED');

	var targetDir = product.ele('Directory', {
		'Id': 'TARGETDIR',
		'Name': 'SourceDir'
	});

	this.programFiles.createElems(targetDir);

	var instExecSeq = product.ele('InstallExecuteSequence');
	instExecSeq.ele('RemoveExistingProducts', {
		'After': 'InstallValidate'
	});

	var feature = product.ele('Feature', {
		'Id': 'Complete',
		'Level': '1'
	});

	this.programFiles.traverseFiles(function (f) {
		feature.ele('ComponentRef', {
			'Id': f.componentId
		});
	});

	return wix.end({ pretty: true, indent: '  ', newline: '\n' });
};

SimpleMsi.prototype.find = function (query) {
	return this.programFiles.find(query);
};

SimpleMsi.prototype.getNextId = function () {
	return "U" + (this.nextId++);
};

SimpleMsi.prototype.downloadWix = function (cacheFolder, onComplete) {
	var stats = fs.lstat(cacheFolder, function (err, stats) {
		if (!err) {
			onComplete(null);
			return;
		}

		mkdirp(cacheFolder, function (err) {
			if (err) {
				onComplete(err);
				return;
			}

			process.stdout.write('downloading wix39-binaries.zip' + '\n');
			httpreq.download('http://download-codeplex.sec.s-msft.com/Download/Release?ProjectName=wix&DownloadId=1421697&FileTime=130661188723230000&Build=20959', 
				cacheFolder + '/wix39-binaries.zip', 
			function (err, prog) {
				process.stdout.write(prog.percentage + '\n');
			}, function (err, res) {
			    if (err) {
			    	onComplete(err);
			    	return;
			    }

			     var r = fs.createReadStream(cacheFolder + '/wix39-binaries.zip').pipe(unzip.Extract({ path: cacheFolder }));
			     r.on('end', function () {
			     	process.stdout.write('extraction complete.\n');
			     	onComplete(null);
			     });
			});
		});
	});
};

SimpleMsi.prototype.build = function (dstPath, onComplete) {
	var wxs = this.toWxs();

	var cacheFolder = './cache/wix';
	this.downloadWix(cacheFolder, function (err) {
		process.stdout.write('download completed\n');
		fs.writeFileSync('./build/setup.wxs', wxs);

		spawnSync(cacheFolder + '/candle.exe', ['./build/setup.wxs', '-out', './build/setup.wixobj']);
		spawnSync(cacheFolder + '/light.exe', ['./build/setup.wixobj', '-out', dstPath]);


		if (onComplete) {
			onComplete(null);
		}
	});
};

var ProgramFilesFolder = function (msi, parent, name) {
	this.msi = msi;
	this.parent = parent;
	this.name = name;
	this.id = msi.getNextId();
	this.folders = {};
	this.files = {};

	if (!parent) {
		this.id = this.name;
		this.name = null;
	}
};

ProgramFilesFolder.prototype.folder = function (name) {
	if (!this.folders.hasOwnProperty(name)) {
		var f = new ProgramFilesFolder(this.msi, this, name);
		this.folders[name] = f;
	}

	return this.folders[name];
};

ProgramFilesFolder.prototype.contents = function (base, glob) {
	var baseFolder = this;

	var paths = globby.sync(glob, {
		'cwd': base
	});

	paths.forEach(function (p) {
		var components = p.split('/');
		var folder = baseFolder;

		for (var i = 0; i < components.length - 1; ++i) {
			folder = folder.folder(components[i]);
		}

		var fileName = components[components.length - 1];

		if (fs.lstatSync(base + '/' + p).isFile()) {
			if (!folder.files.hasOwnProperty(fileName)) {
				folder.files[fileName] = new ProgramFilesFile(folder, fileName, base + '/' + p);
			}
		} else {
			folder.folder(fileName);
		}
	});
};

ProgramFilesFolder.prototype.find = function (query) {

	for (var f in this.files) {
		if (!this.files.hasOwnProperty(f)) {
			continue;
		}

		if (f == query) {
			return this.files[f];
		}
	}

	for (var k in this.folders) {
		if (!this.folders.hasOwnProperty(k)) {
			continue;
		}

		var res = this.folders[k].find(query);
		if (res != null) {
			return res;
		}
	}

	return null;
};

ProgramFilesFolder.prototype.createElems = function (targetDir) {
	var attrs = {};

	if (this.name) {
		attrs['Name'] = this.name;
	}

	if (this.id) {
		attrs['Id'] = this.id;
	}

	var directory = targetDir.ele('Directory', attrs);

	for (var k in this.folders) {
		if (!this.folders.hasOwnProperty(k)) {
			continue;
		}

		this.folders[k].createElems(directory);
	}

	for (var f in this.files) {
		if (!this.files.hasOwnProperty(f)) {
			continue;
		}

		this.files[f].createElem(directory);
	}
}

ProgramFilesFolder.prototype.traverseFiles = function (cb) {
	for (var k in this.folders) {
		if (!this.folders.hasOwnProperty(k)) {
			continue;
		}

		this.folders[k].traverseFiles(cb);
	}

	for (var f in this.files) {
		if (!this.files.hasOwnProperty(f)) {
			continue;
		}

		cb(this.files[f]);
	}
};

var ProgramFilesFile = function (folder, name, srcPath) {
	this.folder = folder;
	this.name = name;
	this.guid = uuid.v4().toUpperCase();
	this.path = path.resolve(srcPath);
	this.fileId = folder.msi.getNextId();
	this.componentId = this.fileId + 'C';
};

ProgramFilesFile.prototype.createElem = function (targetDir) {
	var component = targetDir.ele('Component', {
		'Id': this.componentId,
		'Guid': this.guid,
		'Win64': this.folder.msi.arch == 'x64' ? 'yes' : 'no'
	});

	var file = component.ele('File', {
		'Id': this.fileId,
		'Name': this.name,
		'Source': this.path,
		'KeyPath': 'yes'
	});
};

var StartMenuFolder = function (msi, parent, name) {
	this.msi = msi;
	this.parent = parent;
	this.name = name;
	this.folders = {};
};

StartMenuFolder.prototype.folder = function (name) {
	if (!this.folders.hasOwnProperty(name)) {
		var f = new StartMenuFolder(this.msi, this, name);
		this.folders[name] = f;
	}

	return this.folders[name];
};

StartMenuFolder.prototype.shortcut = function (ref, name) {
	if (typeof ref == 'string') {
		ref = this.msi.find(ref);
	}


};

var Desktop = function (msi) {

};

Desktop.prototype.shortcut = function (ref, name) {

};

module.exports = SimpleMsi;
