function Prompt(data, root) {
    this._data = data;
    this._root = root;
    this._html = null;
    this._context = null;
    this._events = [];
    this._buttonHtml = null;
    this._this = null;
    this.init();
    this.fetch();
};


Prompt.prototype.show = function show() {
    this.render();
    this.ready();
    $('.top-titlebar').children().css('visibility', 'hidden');
};

Prompt.prototype.close = function close() {
    this.destroy();
};

//初始化组件根节点和配置
Prompt.prototype.init = function() {
    if (!this._root) {
        this._root = $('body');
    }
}

//加载 css 和 js 资源
Prompt.prototype.fetch = function() {
    this._html = require(__dirname + '/prompt.html');
    this._buttonHtml = $(this._html).find('#prompt-button-template');

}

//内容渲染
Prompt.prototype.render = function() {
    this._context = {
        title: this._data.title,
        body: this._data.body
    }

    let buttons = [];
    let buttonHtml = this._buttonHtml.html();
    let length = this._data.buttons.length;
    this._data.buttons.forEach(function(val, i) {
        let template = Handlebars.compile(buttonHtml);
        let context = {
            index: i,
            text: val.text,
            lastone: (i == length - 1),
            size: Math.floor(12 / length)
        };

        buttons[i] = template(context);
    });
    this._context.buttons = new Handlebars.SafeString(buttons.join(''));

    let template = Handlebars.compile(this._html);
    this._this = $(template(this._context));

    this._this.css('animation-name', 'animation-in');
    this._this.css('animation-duration', '0.3s');
    this._root.append(this._this);

}

//进行数据绑定等操作
Prompt.prototype.ready = function() {
    if (this._data.buttons) {
        for (var i = 0; i < this._data.buttons.length; i++) {
            this._events[i] = this._data.buttons[i].click;
            this._this.find('#prompt_buttons' + i).click(this._events[i]);
        }
    }
}

//解除所有事件监听，删除所有组件节点
Prompt.prototype.destroy = function() {

    this._this.css('animation-name', 'animation-out');
    this._this.css('animation-duration', '0.3s');
    let destroyObj = this._this;
    this._this.get(0).addEventListener("webkitAnimationEnd", function() {
        //动画结束时事件
        destroyObj.remove();
        $('.top-titlebar').children().css('visibility', 'visible');
    }, false);

    this._this = null;
    this._html = null;
    this._data = null;
    this._context = null;
    this._events = null;
    this._buttonHtml = null;
}
module.exports = Prompt;
