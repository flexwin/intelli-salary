// In renderer process (web page).
const {ipcRenderer} = require('electron')
const shell = require('electron').shell;
const fs = require('fs');
const BrowserWindow = require('electron').remote.BrowserWindow
const path = require('path');
const Store = require('electron-store');
const config = new Store();
const XLSX = require('../lib/xlsx');
const Prompt = require('../desktop/js/prompt/index');

const {
    remote
} = require('electron')
const {
    Menu,
    MenuItem
} = remote

let smtpConfig = {
    "host": "smtp.exmail.qq.com",
    "port": 465,
    "secure": true,
    "auth": {
        "fromName": "工资条（测试）",
        "user": "salary@example-qq.com",
        "pass": "123456"
    },
    "cells": []
};

// 获取配置信息
{
    if (config.get('smtpConfig')) {
        smtpConfig = config.get('smtpConfig');
    }
    config.set('smtpConfig', smtpConfig);
}

require.extensions['.html'] = function (module, filename) {
    module.exports = fs.readFileSync(filename, 'utf8');
};

const MAIL_STATE = {
    RUNNING: 'running',
    PAUSE: 'pause',
    WARN: 'warn',
    SUCCESS: 'success',
    ERROR: 'error',
    WAITING: 'waiting'
};

var scrollHeight;
$(document).ready(() => {
    $('#loadStaffdata').click(function () {
        ipcRenderer.send('open-staff-file-dialog-message', '');
    });

    $('#loadSalarydata').click(function () {
        ipcRenderer.send('open-salary-file-dialog-message', '');
    });

    $('#sendmail').click(function () {
        let small = $('.input-small').val();
        let mailhead = $('#mailhead').val();
        scrollHeight = 0;
        $('body').scrollTop(scrollHeight);
        ipcRenderer.send('sendmail-message', {
            subjectline: small,
            mailhead: mailhead
        });
        $('#loadSalarydata').prop('disabled', true);
        $('#loadStaffdata').prop('disabled', true);
        $(this).prop('disabled', true);
        $('#sendmail_text').text('正在发送中…');
        //config.set('mailhead', mailhead);
    })

    //当前月份－1
    var date = new Date();
    var year = date.getFullYear();
    var month = date.getMonth();
    if (month == 0) {
        year -= 1;
        month = 12;
    }
    var small = `${year}年${month}月工资条`;

    var data = {
        message: ''
    };
    var vm = new Vue({
        el: '#input-small',
        data: data
    });

    vm.$watch('message', function (newVal, oldVal) {
        var mailhead = mailhead = `感谢您对公司做出的贡献和努力，现向您发送${newVal}明细： 
    注：1、个人收入所得不得向他人泄露，亦不得询问本公司其他员工收入所得； 2、对工资清单有疑问的，自收到邮件起三个工作日内向人力资源部提出异议；`;
        $('#mailhead').val(mailhead);
    });
    data.message = small;


    var openAppConfigBtn = $('#openAppConfigBtn');
    openAppConfigBtn.click(randerSettings);

    let aboutUs = $('#aboutUs');
    aboutUs.click(openAboutWindow);

    ipcRenderer.send('ready-message');
});

ipcRenderer.on('salary-ready-reply', (event, dbExists) => {
    if (dbExists) {
        var prompt = new Prompt({
            title: '发现有上次处理的工资条，是否继续发送？',
            body: ``,
            buttons: [{
                text: '放弃，重新加载数据表',
                click: function () {
                    ipcRenderer.send('open-salary-file-dialog-message');
                    prompt.close();
                }
            }, {
                text: '继续处理上次未发送成功的数据',
                click: function () {
                    ipcRenderer.send('salary-get-message');
                    prompt.close();
                }
            }]
        });
        prompt.show();
    } else {
        ipcRenderer.send('open-salary-file-dialog-message');
    }
});

ipcRenderer.on('salary-get-reply', (event, docs) => {
    console.log(docs)
    randerTable(docs);
})

ipcRenderer.on('verifyConfiguration-reply', (event, error) => {
    $('#loadingIcon').addClass('hidden');
    if (error) {
        alert('验证[失败]，请检查是否正确。');
    } else {
        alert('验证[成功]!');
    }
    $('#verifyConfigurationBtn').prop('disabled', false);
})

//导入数据后渲染表格
function randerTable(docs) {
    if (!docs.data) {
        return;
    }

    var template = Handlebars.compile($("#salarylist-template").html());
    var rowHtml = '';

    if (docs.cells.length > 0) {
        smtpConfig.cells = docs.cells;
    }


    docs.data.forEach(function (value, index) {
        rowHtml += '<tr>';
        rowHtml += createRowHtml(value, smtpConfig.cells, index + 1);
        rowHtml += '</tr>';
    });

    var cells = [];

    var i = 0;
    smtpConfig.cells.forEach(function (value, index) {
        if (value.used) {
            cells[i] = value;
        }
        i++;
    });

    var data = {
        'cells': cells,
        'rowHtml': rowHtml
    };

    $('#salarylist').html(template(data));
}

function randerSettings() {
    $('.body-div').css('padding-top', '80px');
    $('#loadSalarydata').prop('disabled', true);
    $('#loadStaffdata').prop('disabled', true);
    $('#sendmail').prop('disabled', true);

    var template = Handlebars.compile($("#salarysettings-template").html());
    var data = smtpConfig;
    $('#subject-form').hide();
    $('.body-div').html(template(data));

    $('#verifyConfigurationBtn').click(function () {
        setSmtpConfig();
        $('#loadingIcon').removeClass('hidden');
        $(this).prop('disabled', true);
        ipcRenderer.send('verifyConfiguration-message', smtpConfig);
    });

    //保存
    $('#save-settings').click(function () {
        setSmtpConfig();

        let fromName = $('#fromName').val();
        let _cells = [];
        $("input[name='cells']").each(function (index) {
            _cells[index] = {
                name: $(this).get(0).value,
                used: $(this).get(0).checked
            };
        });

        smtpConfig.auth.fromName = fromName;
        smtpConfig.cells = _cells;
        config.set('smtpConfig', smtpConfig);
        location.reload();
    });
}

function setSmtpConfig() {
    let host = $('#sendserver').val();
    let port = $('#sendprot').val();
    let user = $('#inputEmail').val();
    let pass = $('#inputPassword').val();


    smtpConfig.host = host;
    smtpConfig.port = Number(port);
    smtpConfig.auth.user = user;
    smtpConfig.auth.pass = pass;
    smtpConfig.secure = true;
}

function createRowHtml(doc, cells, index) {
    var html = `<td style="width:15px">
        <center><i id="state_channel_${doc._id}" class="hidden" aria-hidden="true"></i></center>
    </td>
    `;
    cells.forEach(function (value, index, array) {
        if (value.used) {
            cvalue = doc[value.name]
            if (undefined === cvalue) {
                cvalue = ' '
            }
            html += `<td>${cvalue}</td>`;
        }
    })
    return html;
}

ipcRenderer.on('sendmail-reply', (event, docs) => {
    if (docs.length === 0) {
        $('#sendmail_text').text('发送完成');
        return;
    } else {
        $('#sendmail_text').text('重新发送');
        $('#loadStaffdata').prop('disabled', false);
        $('#sendmail').prop('disabled', false);

    }
    randerTable(docs);
})

//更新进度条通知
ipcRenderer.on('progress-percentage-reply', (event, arg) => {
    if (arg < 1) {
        NProgress.start()
    } else if (arg >= 100) {
        NProgress.done()
    } else {
        NProgress.inc()
    }
});


ipcRenderer.on('update-state-reply', (event, arg) => {
    channelStateManage(arg);
    $('body').scrollTop(scrollHeight += 37);
});


function channelStateManage(doc) {
    var stateIcon = $('#state_channel_' + doc._id);
    stateIcon.removeClass();
    switch (doc.state) {
        case MAIL_STATE.RUNNING:
            stateIcon.addClass('fa fa-refresh fa-spin icon-state-success fa-lg');
            break;
        case MAIL_STATE.WARN:
            stateIcon.addClass('fa fa-exclamation-triangle icon-state-warn fa-lg');
            break;
        case MAIL_STATE.SUCCESS:
            stateIcon.addClass('fa fa-check-circle icon-state-success fa-lg');
            break;
        case MAIL_STATE.ERROR:
            stateIcon.addClass('fa fa-exclamation-triangle icon-state-error fa-lg');
            break;
        case MAIL_STATE.PAUSE:
            stateIcon.addClass('fa fa-pause-circle icon-state-warn fa-lg');
            break;
        case MAIL_STATE.WAITING:
            stateIcon.addClass('fa fa-clock-o icon-state-info fa-lg');
            break;
        default:
            stateIcon.addClass('hidden');
            break;
    }
}

var terminalWin;

function openAboutWindow() {
    if (terminalWin == null) {
        const modalPath = path.join('file://', __dirname, 'about.html')
        terminalWin = new BrowserWindow({
            width: 400,
            height: 280,
            //icon: path.join(__dirname, 'images/terminal.png')
        })
        terminalWin.on('close', function () {
            terminalWin = null;
            //ipcRenderer.send('terminal-window-message', 'close');
        })
        terminalWin.loadURL(modalPath)
    }

    terminalWin.show();
}

function getAsarUnpackedPath(basePath) {
    return `${path.dirname(path.dirname(__dirname))}/app.asar.unpacked/${basePath}`;
}
