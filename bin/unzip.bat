@set @junk=1 /* vim:set ft=javascript:
@echo off
cscript //nologo //e:jscript "%~dpn0.bat" %*
goto :eof
*/

function unzip(file, dir) {
  var fso = new ActiveXObject('Scripting.FileSystemObject');
  if (!fso.FolderExists(dir)) {
    fso.CreateFolder(dir);
  }
  var shell = new ActiveXObject('Shell.Application');
  var dst = shell.NameSpace(fso.getFolder(dir).Path);
  var zip = shell.NameSpace(fso.getFile(file).Path);
  // http://msdn.microsoft.com/en-us/library/ms723207.aspx
  // 4: Do not display a progress dialog box.
  // 16: Click "Yes to All" in any dialog box displayed.
  dst.CopyHere(zip.Items(), 4 + 16);
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
      destdir: null
    }
  };
  var i;
  for (i = 0; i < args.length; ++i) {
    if (args[i].match(/^(-h|--help)$/)) {
      res.option.help = true;
    } else if (args[i].match(/^(-d)$/)) {
      res.option.destdir = args[++i];
    } else if (args[i].match(/^-/)) {
      throw Error('unknown option: ' + args[i]);
    } else {
      res.args.push(args[i]);
    }
  }
  return res;
}

function usage() {
  WScript.Echo('usage: unzip [options] zipfile');
  WScript.Echo('  -h --help    print this text');
  WScript.Echo('  -d exdir     extract files into exdir');
  WScript.Quit(1);
}

function main() {
  var args, zipfile, destdir;
  args = parse_arguments(getargs());
  if (args.option.help || args.args.length === 0) {
    usage();
  }
  zipfile = args.args[0];
  destdir = args.option.destdir;
  if (destdir === null) {
    destdir = ".";
  }
  unzip(zipfile, destdir);
}

main();
