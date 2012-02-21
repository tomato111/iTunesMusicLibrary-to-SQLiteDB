function Ini() {
    this.initialize.apply(this, arguments);
}

Ini.prototype = {
    initialize: function (file) {
        this.clear();
        if (file != null) this.open(file);
    },

    // clearメソッド - 全削除
    clear: function () {
        try { delete this.items; } catch (e) { }
        this.items = new Array();  // 項目
        this.filename = null;    // ファイル名
    },

    // openメソッド - Iniファイルの読込
    open: function (filename) {
        try {
            var fso = new ActiveXObject("Scripting.FileSystemObject");  // FileSystemObjectを作成
            var iniStr = fso.OpenTextFile(filename, 1, false);        // ファイルを開く

            this.clear();
            var sectionname = null;
            var p = -1;

            while (!iniStr.AtEndOfStream) {
                var line = iniStr.ReadLine();
                line = line.replace(/^[ \t]+/, "");  // 先頭の空白は削除
                if (!line.match(/^(?:;|[ \t]*$)/))    // ;で開始しない, 空行ではない
                {
                    if (line.match(/^\[(.+)\][ \t]*$/))  // セクション行
                    {
                        sectionname = RegExp.$1;
                        this.items[sectionname] = new Array();  // セクション行を追加
                    }
                    else if (sectionname != null && (p = line.indexOf('=')) >= 0) {
                        var keyname = line.substr(0, p);
                        var value = line.substr(p + 1, line.length - p - 1);

                        this.items[sectionname][keyname] = value;
                    }
                }
            }

            this.filename = filename;

            iniStr = null; delete iniStr;
            fso = null; delete fso;

            return true;
        }
        catch (e) {
            this.filename = filename;
            return false;
        }
    },

    // updateメソッド - iniファイルの更新
    update: function (filename) {
        if (filename != null) this.filename = filename;

        try {
            var fso = new ActiveXObject("Scripting.FileSystemObject");  // FileSystemObjectを作成
            var ini = fso.OpenTextFile(this.filename, 2, true)

            for (var sectionname in this.items) {
                ini.WriteLine('[' + sectionname + ']');
                for (var keyname in this.items[sectionname])
                    ini.WriteLine(keyname + '=' + this.items[sectionname][keyname]);
                ini.WriteLine('');
            }
            ini = null; delete ini;
            fso = null; delete fso;

            this.open(this.filename);
            return true;
        }
        catch (e) {
            return false;
        }
    },

    // setItemメソッド - 項目の値設定
    setItem: function (sectionname, keyname, value, updateflag) {
        if (updateflag == null) updateflag = true;

        if (!(sectionname in this.items))
            this.items[sectionname] = new Array();

        this.items[sectionname][keyname] = value;

        if (updateflag && this.filename != null) this.update();
    } 
};
