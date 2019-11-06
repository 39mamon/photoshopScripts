 /**
  * author : satoshi takaguchi
  * comment : psdからテキストをズバッと抜き出します
  */

 function TextExport() {
	//改行コードをOSに応じて取得
	if ($.os.search(/windows/i) != -1){
		LINEFEED_CODE = "windows";
	}else{
		LINEFEED_CODE = "macintosh";
	}
	//エラー処理 > 開いているPSDが見当たらない場合
	if (app.documents.length === 0) {
		alert("PSDファイルを開いてから実行して下さい。");
		return;
	}

	//ファイル名と場所を指定するダイアログ
	var saveFile = File.saveDialog("CSVの出力名と場所を指定してください。");
	//エラー処理 > キャンセルした場合、スクリプト中止
	if(saveFile == null){
		return;
	}
	filePath = saveFile.path + "/" + saveFile.name+".csv";
	
	//出力するテキストファイルを作成
	var OUTPUT_FILE	= new File(filePath);
	//エンコード設定
	OUTPUT_FILE.encoding = "Shift-JIS";
	//改行設定
	OUTPUT_FILE.linefeed = LINEFEED_CODE;
	//ファイルを開く
	OUTPUT_FILE.open("w", "TEXT", "????");
	OUTPUT_FILE.writeln("Text"+","+"Font"+","+"Size"+","+"leading"+","+"line-height"+","+"letter-spacing"+","+"HEX"+","+"R"+","+"G"+","+"B");
	//レイヤーの走査結果を書き込んで保存
	ScanAndWrite(app.activeDocument, OUTPUT_FILE, '/');
	OUTPUT_FILE.close();
	alert("CSVへ出力完了しました!!!")
}

/*出力*/
function ScanAndWrite(targetDocument, OUTPUT_FILE, path) {
	var resultText = "";
	var R = 0;
	var G = 0;
	var B = 0;
	//レイヤーを走査
	var layers = targetDocument.layers;
	for (var layerIndex = 0; layerIndex < layers.length; layerIndex++){

		//現在のレイヤーを取得
		var currentLayer = layers[layerIndex];

		//現在のレイヤーがレイヤーセットか否か
		if (currentLayer.typename == "LayerSet") {
			ScanAndWrite(currentLayer, OUTPUT_FILE, path + currentLayer.name + '/');
		} else {
			if ( (currentLayer.visible) && (currentLayer.kind == LayerKind.TEXT) ){
				resultText = "";
				R = Math.round(currentLayer.textItem.color.rgb.red);
				G = Math.round(currentLayer.textItem.color.rgb.green);
				B = Math.round(currentLayer.textItem.color.rgb.blue);
				fontSize = Math.floor(currentLayer.textItem.size*100)/100;
				// TODO : 絶対やるべきではないけど、スマートな解決までは...
				try {
					leading = Math.floor(currentLayer.textItem.leading*100)/100;
					lineHeight = round(leading/fontSize, 2);
				} catch (error) {
					leading = "leading is auto or undefinde or error?";
					lineHeight = "leading is auto or undefinde or error?";					
				}
                letterSpace = currentLayer.textItem.tracking / 1000+"em";

				resultText +=DelLinefeed(currentLayer.textItem.contents)+","+currentLayer.textItem.font+","+fontSize+","+leading+","+lineHeight+","+letterSpace+","+RgbToHex(R,G,B)+","+R+","+G+","+B;
				OUTPUT_FILE.writeln(resultText);
			}
		}
	}
}

/*改行コード、コンマ削除*/
function DelLinefeed(txt){
	var result = "";
	result = txt.replace(/[\n\r]/g,"");
	result = result.replace (/,/g, '');
	return result;
}

/*RBGをHEXに変換*/
function RgbToHex(r,g,b){
	return "#" +("0" + Math.round(r,10).toString(16)).slice(-2)+("0" + Math.round(g,10).toString(16)).slice(-2)+("0" + Math.round(b,10).toString(16)).slice(-2);
}

/*
	MDNの提案されている四捨五入メソッド
	https://developer.mozilla.org/ja/docs/Web/JavaScript/Reference/Global_Objects/Math/round#A_better_solution
*/
function round(number, precision) {
	var shift = function (number, precision, reverseShift) {
	  if (reverseShift) {
		precision = -precision;
	  }  
	  var numArray = ("" + number).split("e");
	  return +(numArray[0] + "e" + (numArray[1] ? (+numArray[1] + precision) : precision));
	};
	return shift(Math.round(shift(number, precision, false)), precision, true);
}

/*実行*/
TextExport();