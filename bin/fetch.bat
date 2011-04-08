@set @junk=1 /* vim:set ft=javascript:
@echo off
cscript //nologo //e:jscript "%~dpn0.bat" %*
goto :eof
*/

// http://msdn.microsoft.com/en-us/library/ms676152.aspx
// SaveOptionsEnum
var adSaveCreateNotExist = 1;
var adSaveCreateOverWrite = 2;

// http://msdn.microsoft.com/en-us/library/ms675277.aspx
// StreamTypeEnum
var adTypeBinary = 1;
var adTypeText = 2;

var USER_AGENT = 'fetch.js/1.0';

function fetch(url, file) {
  var xhr;
  if (url.match(/^https?:/)) {
    // use ServerXMLHTTP for https
    xhr = new ActiveXObject('MSXML2.ServerXMLHTTP');
    // FIXME: required?
    //var SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056;
    //xhr.setOption(2, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS);
  } else {
    // ServerXMLHTTP doesn't support ftp.
    xhr = new ActiveXObject('MSXML2.XMLHTTP.3.0');
  }
  xhr.open('GET', url, false);
  xhr.setRequestHeader('User-Agent', USER_AGENT);
  xhr.send();
  if (xhr.readyState !== 4) {
    throw new Error('REQUEST INCOMPLETE');
  }
  if (url.match(/^https?/) && xhr.status !== 200) {
    throw new Error('HTTP STATUS: ' + xhr.status);
  }

  var strm = new ActiveXObject('ADODB.Stream');
  strm.Type = adTypeBinary;
  strm.Open();
  strm.Write(xhr.responseBody);
  strm.SaveToFile(file, adSaveCreateOverWrite);
  strm.Close();
}

function filename(url) {
  url = url.replace(/[?#].*$/, '');
  if (url.match(/\/$/)) {
    return 'index.html';
  }
  return url.match(/[^\/]+$/);
}

function getargs() {
  var args = [], i;
  for (i = 0; i < WScript.Arguments.length; ++i) {
    args.push(WScript.Arguments(i));
  }
  return args;
}

function parse_arguments(args) {
  var res = {
    args: [],
    option: {
      help: false,
      destfile: null
    }
  };
  var i;
  for (i = 0; i < args.length; ++i) {
    if (args[i].match(/^(-h|--help)$/)) {
      res.option.help = true;
    } else if (args[i].match(/^(-o)$/)) {
      res.option.destfile = args[++i];
    } else if (args[i].match(/^-/)) {
      throw Error('unknown option: ' + args[i]);
    } else {
      res.args.push(args[i]);
    }
  }
  return res;
}

function usage() {
  WScript.Echo('usage: fetch [options] url');
  WScript.Echo('  -h --help    print this text');
  WScript.Echo('  -o           write documents to FILE.');
  WScript.Quit(1);
}

function main() {
  var args, url, destfile;
  args = parse_arguments(getargs());
  if (args.option.help || args.args.length === 0) {
    usage();
  }
  url = args.args[0];
  destfile = args.option.destfile;
  if (destfile === null) {
    destfile = filename(url);
  }
  fetch(url, destfile);
}

main();
