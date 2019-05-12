const {
    app
} = require('electron');
const log = require('electron-log');

const XLSX = require('xlsx');
const Datastore = require('nedb');
const fs = require('fs');

let db = null;

function loaddata() {
    db = new Datastore({
        filename: getDatabasePath(),
        timestampData: true,
        autoload: true
    });
    db.loadDatabase()
}

function read(path) {
    if (db === null)
        loaddata();

    var workbook = XLSX.readFile(path, {
        sheetRows: 0
    });
    var first_sheet_name = workbook.SheetNames[0];
    var result = {};
    /* Get worksheet */
    var worksheet = workbook.Sheets[first_sheet_name];
    result = XLSX.utils.sheet_to_json(worksheet);
    return result
}

function readOnly(path) {
    var workbook = XLSX.readFile(path, {
        sheetRows: 0
    });
    var first_sheet_name = workbook.SheetNames[0];
    var result = {};
    /* Get worksheet */
    var worksheet = workbook.Sheets[first_sheet_name];
    result = XLSX.utils.sheet_to_json(worksheet);
    return result
}

function save(doc, callback) {
    if (db === null)
        loaddata();
    // for (var i in doc) {
    //     doc[i]._index = parseInt(Date.now() + i)
    //     doc[i]._state = 0
    // }
    // Callback is optional
    db.insert(doc, function (err, newDoc) {
        callback(err, newDoc);
        db.loadDatabase()
    });
}

function update(query, update, options, callback) {
    if (db === null)
        loaddata();
    db.update(query, update, options, callback);
    db.loadDatabase();
}

function get(searchDoc, callback) {
    if (db === null)
        loaddata();
    if (typeof (searchDoc) === 'function') {
        callback = searchDoc;
        searchDoc = {};
    }
    db.loadDatabase();
    db.find(searchDoc).sort({
        '_index': 1
    }).exec(callback);
}

function getDatabasePath() {
    return `${app.getPath('userData')}/db/salary.db `;
}

exports.read = read;
exports.readOnly = readOnly;
exports.update = update;
exports.save = save;
exports.get = get;
exports.getDatabasePath = getDatabasePath;
