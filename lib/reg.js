// RFC 5322 http://emailregex.com/
const mail_reg = /^(([^<>()\[\]\\.,:\s@"]+(\.[^<>()\[\]\\.,:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
// 18位身份证
const idcard_reg = /[1-9]\d{5}[1-2]\d{3}((0\d)|(1[0-2]))(([0|1|2]\d)|3[0-1])\d{3}(\d|X|x)/;

/**
 * 是否邮箱
 * @param str
 * @returns {boolean}
 */
function isEmail(str) {
    if (str === null) {
        return false
    } else {
        return mail_reg.test(str)
    }
}

/**
 * 是否身份证
 * @param str
 * @returns {boolean}
 */
function isIdcard(str) {
    if (str === null) {
        return false
    } else {
        return idcard_reg.test(str)
    }
}

exports.isEmail = isEmail;
exports.isIdcard = isIdcard;
