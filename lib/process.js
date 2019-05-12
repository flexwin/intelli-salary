const {
    app,
    ipcMain,
    autoUpdater,
    dialog
} = require('electron');
const fs = require('fs');
require.extensions['.html'] = function (module, filename) {
    module.exports = fs.readFileSync(filename, 'utf8')
};

var queue = require('queue');
const log = require('electron-log');
const Store = require('electron-store');
const config = new Store();
const XLSX = require('../lib/xlsx');
const reg = require('../lib/reg');
const mail = require('../lib/mail');

const MAIL_STATE = {
    RUNNING: 'running',
    PAUSE: 'pause',
    WARN: 'warn',
    SUCCESS: 'success',
    ERROR: 'error',
    WAITING: 'waiting'
};

//初始化完成后返回
ipcMain.on('ready-message', (event, arg) => {
    let salary_db_path = XLSX.getDatabasePath();
    event.sender.send('salary-ready-reply', fs.existsSync(salary_db_path))
});

//读取的excel
ipcMain.on('salary-get-message', (event, arg) => {
    var query = {
        _state: {
            $ne: 'success'
        }
    };
    XLSX.get(query, (err, docs) => {
        event.sender.send('salary-get-reply', {
            'cells': [],
            'data': docs
        })
    });
});

ipcMain.on('app-message', (event, arg) => {
    event.sender.send('app-reply', app.getVersion());
});


var q = queue({
    'concurrency': 1
});

// 选择员工工资表
ipcMain.on('open-salary-file-dialog-message', (event, arg) => {
    dialog.showOpenDialog({
        title: '请选择员工工资表',
        filters: [{
            name: 'xlsx',
            extensions: ['xls', 'xlsx']
        }],
        properties: ['openFile']
    }, function (files) {
        if (files) {
            var data = XLSX.read(files[0]);
            // 覆盖
            if (fs.existsSync(XLSX.getDatabasePath())) {
                fs.unlinkSync(XLSX.getDatabasePath());
            }
            XLSX.save(data, (err, newDoc) => {
                var keys = Object.keys(data[0]);
                var values = Object.values(data[1]);
                var i = 0;
                var _cells = [];
                for (var i = 0; i < keys.length; i++) {
                    let _used = getCellUsed(keys[i]);
                    // 找到身份证列
                    let _unique = reg.isIdcard(values[i]);
                    _cells[i] = {
                        'name': keys[i],
                        'used': _used,
                        'unique': _unique
                    };
                }
                // 邮箱固定显示
                _cells[keys.length] = {
                    'name': '邮箱',
                    'used': true
                };
                //保存最新的
                var smtpConfig = config.get('smtpConfig');
                smtpConfig.cells = _cells;
                config.set('smtpConfig', smtpConfig);
                event.sender.send('salary-get-reply', {
                    'cells': _cells,
                    'data': newDoc
                });
            })
        }
    });
});

// 选择员工信息表格
ipcMain.on('open-staff-file-dialog-message', (event, arg) => {
    let salary_db_path = XLSX.getDatabasePath();
    if (!fs.existsSync(salary_db_path)) {
        dialog.showErrorBox('错误提示', '请先选择工资表并确保工资表中员工身份证信息完整');
        return
    }
    dialog.showOpenDialog({
        title: '请选择员工信息表格',
        filters: [{
            name: 'xlsx',
            extensions: ['xls', 'xlsx']
        }],
        properties: ['openFile']
    }, function (files) {
        if (files) {
            var data = XLSX.readOnly(files[0]);
            var keys = Object.keys(data[0]);
            let _isEmail = false;
            let _isIdCard = false;
            let _email = '';
            let _idCard = '';
            data.forEach((item, index, array) => {
                keys.forEach((key) => {
                        _value = item[key];
                        if (reg.isEmail(_value)) {
                            _isEmail = true;
                            _email = _value
                        }
                        if (reg.isIdcard(_value)) {
                            _isIdCard = true;
                            _idCard = _value
                        }
                    }
                );
                log.info("同时存在身份证和邮箱列则更新工资表？", _isIdCard, _isEmail);
                // 同时存在身份证和邮箱列则更新工资表
                if (_isIdCard && _isEmail) {
                    let _unique = getUniqueCellName();
                    let query = {};
                    query[_unique] = _idCard;
                    XLSX.get(query, (err, docs) => {
                        if (err) {
                            log.warn('查询出错', err);
                            return;
                        }
                        docs.forEach(function (employee) {
                            XLSX.update({
                                _id: employee._id
                            }, {
                                $set: {
                                    '邮箱': _email
                                }
                            }, {}, () => {
                                XLSX.get((err, docs) => {
                                    if (err) {
                                        log.warn('查询出错', err);
                                        return;
                                    }
                                    event.sender.send('salary-get-reply', {
                                        'cells': config.get('smtpConfig').cells,
                                        'data': docs
                                    });
                                })
                            })
                        });
                    })
                }
            })
        }
    });
});

/**
 * 获取config中的cell字段是否是使用的（used）
 * @param cellName
 * @returns {boolean}
 */
function getCellUsed(cellName) {
    var smtpConfig = config.get('smtpConfig');
    let _used = true;
    if (smtpConfig.cells) {
        for (var i = 0; i < smtpConfig.cells.length; i++) {
            let _cell = smtpConfig.cells[i];
            if (_cell.name === cellName) {
                _used = _cell.used;
                break;
            }
        }
    }
    return _used;
}

/**
 * 获取config中的cell字段是身份证的cellName
 * @returns {string}
 */
function getUniqueCellName() {
    var smtpConfig = config.get('smtpConfig');
    let _name = '身份证号';
    if (smtpConfig.cells) {
        for (var i = 0; i < smtpConfig.cells.length; i++) {
            let _cell = smtpConfig.cells[i];
            if (_cell.unique === true) {
                _name = _cell.name;
                break;
            }
        }
    }
    return _name;
}

/**
 * 上传excel
 */
ipcMain.on('xlsxuplaod-message', (event, arg) => {
});

/**
 * 验证邮件配置
 */
ipcMain.on('verifyConfiguration-message', (event, arg) => {
    mail.verifyConfiguration(arg, function (error) {
        event.sender.send('verifyConfiguration-reply', error);
    });
});

/**
 * 发送邮件
 */
ipcMain.on('sendmail-message', (event, arg) => {
    console.log("发送邮件：" + arg.subjectline);
    sendmailMessage(event, arg);
});

/**
 * 发送邮件内容
 * @param event
 * @param arg
 */
function sendmailMessage(event, arg) {
    XLSX.get((err, docs) => {
        if (err) {
            dialog.showErrorBox('错误提示', '加载数据有误' + err);
            return;
        }
        event.sender.send('progress-percentage-reply', 0);
        docs.forEach(function (employee) {
            employee.state = MAIL_STATE.WAITING;
            event.sender.send('update-state-reply', employee);
            q.push(function (done) {
                sendMail(employee, event, arg, done)
            })
        });

        q.start(function (err) {
            event.sender.send('progress-percentage-reply', 100);
            log.info('全部处理完毕');
            //event.sender.send('end-sendmail-reply', '');
            var query = {
                _state: {
                    $ne: 'success'
                }
            };
            XLSX.get(query, (err, docs) => {
                let failSize = docs.length;
                let options = {
                    title: '信息',
                    buttons: ['OK'],
                    message: '发送完成'
                };

                if (failSize === 0) {
                    options.detail = '全部发送成功!';

                    //清空数据
                    if (fs.existsSync(XLSX.getDatabasePath()))
                        fs.unlinkSync(XLSX.getDatabasePath());

                    dialog.showMessageBox(options, function (optional) {
                        if (optional === 0) {
                            event.sender.send('sendmail-reply', docs);
                            confirmQuit();
                        }

                    });
                } else {
                    options.detail = `共${failSize}条发送失败，请点击重新发送`;
                    options.buttons = ['重新发送', '取消'];

                    log.warn('发送失败');
                    dialog.showMessageBox(options, function (optional) {
                        if (optional === 0) {
                            sendmailMessage(event, arg);
                        } else if (optional === 1) {
                            event.sender.send('sendmail-reply', docs);
                        }
                    });
                }

            });
        });
    });
}

/**
 * 退出确认
 */
function confirmQuit() {
    let options = {
        title: '信息',
        message: '关闭应用',
        detail: '邮件都发送完毕，是否关闭应用？',
        buttons: ['关闭', '取消'],
    };
    dialog.showMessageBox(options, function (optional) {
        if (optional === 0) {
            app.quit();
        }
    });
}

/**
 * 发送邮件
 * @param employee
 * @param event
 * @param mailOptions
 * @param done
 */
function sendMail(employee, event, mailOptions, done) {
    if (!employee["邮箱"]) {
        log.warn('没有[邮箱]列');
        employee.state = MAIL_STATE.WARN;
        updateMailSate(employee, function () {
            event.sender.send('update-state-reply', employee);
            event.sender.send('progress-percentage-reply', 80);
            done()
        });
        return;
    }
    console.log("接收者:" + employee["_id"] + "（" + employee["邮箱"] + "）,主题:" + mailOptions.subjectline);
    mail.sendMail(employee, mailOptions, function (error, info) {
        if (error) {
            log.error(`[error]:${error}`);
            employee.state = MAIL_STATE.ERROR;
        } else {
            employee.state = MAIL_STATE.SUCCESS;
        }
        updateMailSate(employee, function () {
            event.sender.send('update-state-reply', employee);
            event.sender.send('progress-percentage-reply', 80);
            done();
        });
    });
}

/**
 * 更新工资 表邮件发送状态
 * @param employee
 * @param callback
 */
function updateMailSate(employee, callback) {
    XLSX.update({
        _id: employee._id
    }, {
        $set: {
            _state: employee.state
        }
    }, {}, callback)
}

/**
 * 检查有更新
 */
autoUpdater.on('update-available', function () {
    const options = {
        type: 'none',
        title: '检查更新',
        message: "发现有新版本，正在后台下载",
        buttons: ['OK']
    };

    dialog.showMessageBox(options, function () {

    });
});

/**
 * 检查没有更新
 */
autoUpdater.on('update-not-available', function () {
    const options = {
        type: 'none',
        title: '检查更新',
        message: "您当前已经是最新版！",
        buttons: ['OK']
    };

    dialog.showMessageBox(options, function () {

    });
});

/**
 * 更新下载完成
 */
autoUpdater.on('update-downloaded', function (event, releaseNotes, releaseName, releaseDate, updateUrl, quitAndUpdate) {
    const options = {
        type: 'none',
        title: '下载完毕',
        message: "下载完毕，正在安装并重新打开",
        buttons: ['OK']
    };

    dialog.showMessageBox(options);
    autoUpdater.quitAndInstall();
});

/**
 * 检查远程版本
 */
function checkForUpdates() {
    if (!fs.existsSync('squirrel.exe'))
        return;

    // autoUpdater.setFeedURL('http://intelli-salary.feed.sahara.com');
    // autoUpdater.checkForUpdates();

    // fs.readdir('../', function(err, files) {
    //     if (err) {
    //         log.error(err);
    //         return;
    //     }
    //     var count = files.length;
    //     var results = {};
    //     files.forEach(function(filename) {
    //         if (filename.startsWith('app-') && filename !== `app-${app.getVersion()}`) {
    //             //删除老版本
    //             fs.rmdirAsync(`../${filename}`, function(err, stdout, stderr) {
    //                 if (err != null) {
    //                     console.log('clean build Folder error:' + err);
    //                     return;
    //                 }
    //             });
    //         }
    //     });
    // });
}
