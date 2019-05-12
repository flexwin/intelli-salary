const nodemailer = require('nodemailer');
var Handlebars = require('handlebars');
const log = require('electron-log');
const Store = require('electron-store');
const config = new Store();

var smtpConfig = null;

// create reusable transporter object using the default SMTP transport
var transporter = null;

function sendMail(employee, mailOptions, callback) {
    if (smtpConfig == null) {
        smtpConfig = config.get('smtpConfig');
        if (smtpConfig != null) {
            transporter = nodemailer.createTransport(config.get("smtpConfig"));
        } else {
            log.error('transporter create error');
            return;
        }
    }

    var cellsKey = [];
    var cellsVer = [];
    var i = 0;

    smtpConfig.cells.forEach(function (value) {
        if (value.used) {
            cellsKey[i] = value.name;
            cellsVer[i] = employee[value.name];
            i++;
        }
    });

    var data = {
        'mailhead': mailOptions.mailhead,
        'titles': cellsKey,
        'cells': cellsVer
    };

    var html = require('../desktop/mailtemplate');
    var template = Handlebars.compile(html);
    var mailContent = template(data);

    var _mailOptions = {
        from: `"${smtpConfig.auth.fromName}" ${smtpConfig.auth.user}`, // sender address
        to: employee["邮箱"], // list of receivers
        subject: mailOptions.subjectline, // Subject line
        html: mailContent // html body
    };
    console.log("subject:" + mailOptions.subjectline);

    // send mail with defined transport object
    transporter.sendMail(_mailOptions, callback);
    //callback(_mailOptions, callback);
}

function verifyConfiguration(vsmtpConfig, cb) {
    let verifyTransporter = nodemailer.createTransport(vsmtpConfig);
    verifyTransporter.verify(function (error, success) {
        if (error) {
            log.error(error);
            cb(error);
        } else {
            log.info('Server is ready to take our messages');
            cb();
        }
    });
}

function writeMail() {
}

//Interface
exports.sendMail = sendMail;
exports.verifyConfiguration = verifyConfiguration;
