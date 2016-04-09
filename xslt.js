// '16.04.09 新規作成
// https://msdn.microsoft.com/ja-jp/library/ms762796(v=vs.85).aspx よりコピー
// 上記サンプルスクリプトからの変更箇所: 引数については拡張子を含めたファイル名を受け取る。
var oArgs = WScript.Arguments;

if (oArgs.length == 0)
{
	WScript.Echo ("Usage : cscript xslt.js xml xsl");
	WScript.Quit();
}
xmlFile = oArgs(0);
xslFile = oArgs(1);

var xsl = new ActiveXObject("MSXML2.DOMDOCUMENT.6.0");
var xml = new ActiveXObject("MSXML2.DOMDocument.6.0");
xml.validateOnParse = false;
xml.async = false;
xml.load(xmlFile);

if (xml.parseError.errorCode != 0)
	WScript.Echo ("XML Parse Error : " + xml.parseError.reason);

xsl.async = false;
xsl.load(xslFile);

if (xsl.parseError.errorCode != 0)
	WScript.Echo ("XSL Parse Error : " + xsl.parseError.reason);

try
{
	WScript.Echo (xml.transformNode(xsl.documentElement));
}
catch(err)
{
	WScript.Echo ("Transformation Error : " + err.number + "*" + err.description);
}
