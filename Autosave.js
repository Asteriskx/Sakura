// Autosave for sakura editor
!function(setting){
	var BTN = {
		OK: 1,
		CANCEL: 2
	}
	var CHARSET = {
		SJIS: 0,
		JIS: 1,
		EUC: 2,
		Unicode: 3,
		"UTF-8": 4,
		"UTF-7": 5,
		"Unicodde(BigEndian)": 6
	}
	var LINEBREAK = {
		NOCHANGE: 0,
		CRLF: 1,
		LF: 2,
		CR: 3
	}

	// 保存済みには何もしない
	if (Editor.GetFilename() != "") return

	var fso = new ActiveXObject("Scripting.FileSystemObject")
	var wsh_shell = new ActiveXObject("WScript.Shell")

	var defaults = {
		dir: getDir() + "\\autosave",
		filename_template: "{{number}}.{{MM}}{{DD}}-{{HH}}{{mm}}.txt",
		charset: "UTF-8",
		linebreak: "LF"
	}
	
	var dir = setting.dir || defaults.dir
	var filename_template = setting.filename_template || defaults.filename_template
	var charset = CHARSET[setting.charset || defaults.charset]
	var linebreak = LINEBREAK[setting.linebreak || defaults.linebreak]
	
	var number = Editor.ExpandParameter("$n")
	var filename = createFromTemplate(filename_template, getNowForTemplate(), {number: number})
	
	var path = dir.replace(/\\$/, "") + "\\" + filename.replace(/^\\/, "")
	
	// フォルダ準備
    if (!fso.FolderExists(dir)) 
    {
		fso.CreateFolder(dir)
    }
     else if (fso.GetFolder(dir).Files.Count > 0 && number === "1") {
		// 最初の番号でファイルがフォルダに存在するなら
		var message = "前回起動時のデータが残っています。全削除しますか？？"
		var is_ok = OkCancelBox(message) === BTN.OK
		if (is_ok) {
			fso.DeleteFolder(dir, true)
			fso.CreateFolder(dir)
		} else {
			return
		}
	}
	
	// 保存
	Editor.FileSaveAs(path, charset, linebreak)
	
	// "Number is {{number}}." と {number: 100} から "Number is 100." を作る
	function createFromTemplate(tstr /*, ...args*/){
		var obj = mergeObjects(Array.prototype.slice.call(arguments, 1))
		return tstr.replace(/\{\{(.*?)\}\}/g, function(_, name){
			return obj[trim(name)] || ""
		})
	}
	
	// オブジェクトのプロパティをまとめる
	function mergeObjects(objects){
		var result = {}
        for (var i = 0; i < objects.length; i++)
        {
			var obj = objects[i]
            for (var k in obj)
            {
                if (obj.hasOwnProperty(k))
					result[k] = obj[k]	
			}
		}
		return result
	}
	
	// 文字列前後の空白を削除 .trim() の代わり
	function trim(str){
		return str.replace(/^\s*(.+?)\s*$/, "$1")
	}

	// 指定サイズになるように左側に文字を繰り返しくっつける
	function padLeft(value, len, char){
		char = char || "0"
		var padding = Array.apply(null, Array(len + 1)).join(char)
		return (padding + value).slice(-len)
	}

	// 現在時刻をテンプレート用のオブジェクトフォーマットで取得する
	function getNowForTemplate(){
		var date = new Date()
        return {
			YYYY: padLeft(date.getFullYear(), 4),
			MM: padLeft(date.getMonth() + 1, 2),
			DD: padLeft(date.getDate(), 2),
			HH: padLeft(date.getHours(), 2),
			mm: padLeft(date.getMinutes(), 2),
			ss: padLeft(date.getSeconds(), 2),
			fff: padLeft(date.getMilliseconds(), 3)
		}
	}
	
	// %APPDATA%\sakura\autosave か Documents\sakura\autosave 
	function getDir(){
		var sakura = "\\sakura"
		
		var appdata_sakura = wsh_shell.expandEnvironmentStrings("%AppData%") + sakura
		if (fso.FolderExists(appdata_sakura)) {
			return appdata_sakura
		}
		
		var document_sakura = wsh_shell.specialFolders("MyDocuments") + sakura
		if (!fso.FolderExists(document_sakura)) {
			fso.CreateFolder(document_sakura)
		}
		return appdata_sakura
	}
	
}({
	// 文字列：
	// 「{{number}}」が連番に、「{{YYYY}}」「{{MM}}」、「{{DD}}」、「{{HH}}」、「{{mm}}」、「{{ss}}」、「{{fff}}}」
	// がそれぞれ年月日時分秒ミリ秒になります
	filename_template: null,
	// 文字列：
	// 保存先のフォルダのパス（\が2つ必要です）
	directory: null,
	// 文字列：
	// 「SJIS」、「JIS」、「EUC」、「Unicode」、「UTF-8」、「UTF-7」、「Unicode(BigEndian)」から選びます
	charset: null,
	// 文字列：
	// 「CRLF」、「LF」、「CR」から選びます
	linebreak: null
})