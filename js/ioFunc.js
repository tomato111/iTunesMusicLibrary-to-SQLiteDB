var ioFunc = new function () {

    this.createObject = function (name) {
        return WScript.CreateObject(name);
    };

    this.wshell = this.createObject("WScript.Shell"); 		// WScript�̃V�F��
    this.fs = this.createObject("Scripting.FileSystemObject"); // �t�@�C�� �V�X�e�� �I�u�W�F�N�g
    this.rootPath = WScript.ScriptFullName; 					// ���s�t�@�C���̃p�X
    this.rootDir = this.fs.GetFile(this.rootPath).ParentFolder + "\\"; // ���[�g �f�B���N�g��

    // DB, ADODB
    var dbname = "customdb_sqlite.db";
    var dbConn = new ActiveXObject("ADODB.Connection");
    var dbReco = new ActiveXObject("ADODB.Recordset");
    dbConn.ConnectionString = "DRIVER=SQLite3 ODBC Driver;Database=" + this.rootDir + dbname;
    // INI
    var iniPath = this.rootDir + "setting.ini";
    var ini = new Ini(iniPath);

    // �f�t�H���gini �̍쐬
    this.createINI = function () {
        if (!this.fs.FileExists(iniPath)) {
            ini.setItem("key", "key1", "Artist", false);
            ini.setItem("key", "key2", "Name", false);
            ini.setItem("key", ";key3", "Album", false);
            ini.setItem("key", ";key4", "Track Number", false);
            ini.setItem("field name", "Play Count", "PLAY_COUNT_CD", false);
            ini.setItem("field name", "Rating", "RATING_CD", false);
            ini.setItem("field name", "Play Date UTC", "LAST_PLAYED_CD", false);
            ini.setItem("field name", ";Date Added", "ADDED_CD", false);
            ini.update();
        }
    };

    // �t�@�C���������I�����Ď擾
    this.readFile = function (path) {
        var stream = this.fs.OpenTextFile(path, 1, false, -1);
        var rdStr = [];
        var i = -1;
        var word = "Kind";

        for (var section in ini.items)
            if (section == "key")
                for (var keyname in ini.items["key"])
                    word += "|" + ini.items["key"][keyname];
            else
                for (keyname in ini.items[section])
                    word += "|" + keyname;

        var searchStr = new RegExp(">(" + word + ")<\/.*>(.*)<", "");

        while (!stream.AtEndOfStream) {
            if (stream.ReadLine().match(searchStr)) {
                if (RegExp.$1 == "Name") {
                    i++;
                    rdStr[i] = [];
                }
                rdStr[i][RegExp.$1] = decNumRefToString(RegExp.$2);
            }
        }

        stream.Close();
        this.debugXML(rdStr);
        return rdStr;

        function decNumRefToString(decNumRef) { // ���l�����Q��(10�i��)�𕶎���ɕϊ�
            return decNumRef.replace(/&#(\d+);/ig, function (match, $1, idx, all) {
                return String.fromCharCode($1);
            });
        }
    };

    // �f�[�^�x�[�X�̍쐬
    this.createDB = function () {
        if (!this.fs.FileExists(this.rootDir + dbname)) {
            var createSql = "CREATE TABLE IF NOT EXISTS database_version (ver INTEGER);"
					+ "INSERT INTO database_version VALUES (1);"
					+ "CREATE TABLE IF NOT EXISTS quicktag (url TEXT, subsong INTEGER DEFAULT -1, fieldname TEXT, value TEXT);";

            this.fs.CreateTextFile(this.rootDir + dbname, false);
            dbConn.open();
            dbConn.Execute(createSql);
            dbConn.close();
        }
    };

    //�f�[�^�x�[�X�֑}��
    this.insertDB = function (xmlstr) {
        var fieldname, value, xmlArray, i = 0, keystr = "";

        dbConn.open();

        for (var keyname in ini.items["key"]) {  // �L�[�����񐶐�
            if (i != 0)
                keystr += '+\",\"+';
            keystr += "xmlArray[\"" + ini.items["key"][keyname] + "\"]";
            i++;
        }

        for (i = 0; i < xmlstr.length; i++) {
            if (!xmlstr[i]["Kind"] || !xmlstr[i]["Kind"].match(/(?:�I�[�f�B�I|Audio)/i))
                continue;

            xmlArray = xmlstr[i];

            var key = eval(keystr).replace(/'/g, "''").replace(/undefined/g, "?");  // �L�[���������

            if ("Play Count" in xmlArray) {
                fieldname = ini.items["field name"]["Play Count"];
                value = xmlArray["Play Count"];
                insert();
            }
            if ("Rating" in xmlArray) {
                fieldname = ini.items["field name"]["Rating"];
                value = xmlArray["Rating"] / 20;
                insert();
            }
            if ("Play Date UTC" in xmlArray) {
                fieldname = ini.items["field name"]["Play Date UTC"];
                value = UTCtoLocal(xmlArray["Play Date UTC"]);
                insert();
            }
            if ("Date Added" in xmlArray) {
                fieldname = ini.items["field name"]["Date Added"];
                value = UTCtoLocal(xmlArray["Date Added"]);
                insert();
            }
        }

        ini.clear();
        dbConn.close();
        this.debugDB();

        function insert() {  // �C���T�[�g�y�уA�b�v�f�[�g�֐�
            var searchSql = "SELECT value FROM quicktag WHERE url='" + key + "' AND fieldname='" + fieldname + "';";
            var insertSql = "INSERT INTO quicktag(url, fieldname, value) VALUES ('" + key + "','" + fieldname + "','" + value + "');";
            var updateSql = "UPDATE quicktag SET value='" + value + "' WHERE url='" + key + "' AND fieldname='" + fieldname + "';";
            dbReco.open(searchSql, dbConn);
            if (dbReco.BOF)
                dbConn.Execute(insertSql);
            else if (Str2Int(value) > Str2Int(dbReco.Fields(0).value))
                dbConn.Execute(updateSql);
            dbReco.close();
        }

        function Str2Int(str) {  //��r�\���l�ւ̕ϊ��֐�
            var s = String(str);
            if (s.match(/-/))
                s = s.replace(/[-: ]/g, "");
            return Number(s);
        }

        function doubleFig(num) {  // �����֐�
            if (num < 10)
                num = "0" + num;
            return num;
        }

        function UTCtoLocal(date) {  // UTCdate��Local��,�y�уt�H�[�}�b�g�ϊ��֐�
            var dd = date.replace(/\D/g, "-").split("-");
            var d = new Date(dd[0], dd[1] - 1, dd[2], dd[3], dd[4], dd[5]);
            d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
            return d.getFullYear() + "-" + doubleFig(d.getMonth() + 1) + "-" + doubleFig(d.getDate()) + " "
					+ doubleFig(d.getHours()) + ":" + doubleFig(d.getMinutes()) + ":" + doubleFig(d.getSeconds());
        }
    };

    this.debugDB = function () {
        if (!main.debug) return;
        dbConn.open();
        dbReco.open("select * from quicktag", dbConn);

        var str = "     " + dbReco.Fields(0).Name
				+ "                 " + dbReco.Fields(1).Name
				+ "           " + dbReco.Fields(2).Name
				+ "        " + dbReco.Fields(3).Name + "\n";

        while (!dbReco.EOF) {
            str += dbReco.Fields(0).value.toString()
					+ "\t" + dbReco.Fields(1).value
					+ "\t" + dbReco.Fields(2).value
					+ "\t" + dbReco.Fields(3).value + "\n";
            dbReco.MoveNext;
        }

        dbReco.close();
        dbReco = null;
        dbConn.close();
        dbConn = null;

        WScript.Echo(str);
    };

    this.debugXML = function (xmlstr) {
        if (!init.debug) return;
        var str = "";
        for (var i = 0; i < xmlstr.length; i++) {
            str += "Item" + (i + 1) + "\n";
            for (var key in xmlstr[i]) {
                str += "  " + key + ": " + xmlstr[i][key] + "\n";
            }
        }
        WScript.Echo(str);
    };
};
