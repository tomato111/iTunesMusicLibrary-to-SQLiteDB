function Ini() {
    this.initialize.apply(this, arguments);
}

Ini.prototype = {
    initialize: function (file) {
        this.clear();
        if (file != null) this.open(file);
    },

    // clear���\�b�h - �S�폜
    clear: function () {
        try { delete this.items; } catch (e) { }
        this.items = new Array();  // ����
        this.filename = null;    // �t�@�C����
    },

    // open���\�b�h - Ini�t�@�C���̓Ǎ�
    open: function (filename) {
        try {
            var fso = new ActiveXObject("Scripting.FileSystemObject");  // FileSystemObject���쐬
            var iniStr = fso.OpenTextFile(filename, 1, false);        // �t�@�C�����J��

            this.clear();
            var sectionname = null;
            var p = -1;

            while (!iniStr.AtEndOfStream) {
                var line = iniStr.ReadLine();
                line = line.replace(/^[ \t]+/, "");  // �擪�̋󔒂͍폜
                if (!line.match(/^(?:;|[ \t]*$)/))    // ;�ŊJ�n���Ȃ�, ��s�ł͂Ȃ�
                {
                    if (line.match(/^\[(.+)\][ \t]*$/))  // �Z�N�V�����s
                    {
                        sectionname = RegExp.$1;
                        this.items[sectionname] = new Array();  // �Z�N�V�����s��ǉ�
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

    // update���\�b�h - ini�t�@�C���̍X�V
    update: function (filename) {
        if (filename != null) this.filename = filename;

        try {
            var fso = new ActiveXObject("Scripting.FileSystemObject");  // FileSystemObject���쐬
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

    // setItem���\�b�h - ���ڂ̒l�ݒ�
    setItem: function (sectionname, keyname, value, updateflag) {
        if (updateflag == null) updateflag = true;

        if (!(sectionname in this.items))
            this.items[sectionname] = new Array();

        this.items[sectionname][keyname] = value;

        if (updateflag && this.filename != null) this.update();
    } 
};
