@set @junk=1 /* vim:set ft=javascript:
@echo off
cscript //nologo //e:jscript "%~dpn0.bat" %*
goto :eof
*/

function zip(zipfile, files) {
  var fso = new ActiveXObject('Scripting.FileSystemObject');
  var shell = new ActiveXObject('Shell.Application');

  var process_id = get_process_id();

  // create empty zip (right click -> new file -> compressed (zipped) folder)
  var ozip = fso.CreateTextFile(zipfile, true);
  ozip.Write('PK\x05\x06\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00');
  ozip.Close();

  // copy to zip (drop files to zip folder)
  var zipfolder = shell.NameSpace(fso.GetAbsolutePathName(zipfile));
  var i, j;
  for (i = 0; i < files.length; ++i) {
    var oldcount = thread_count(process_id);
    var f = fso.GetFile(zipfile);
    var oldsize = f.Size;
    // FIXME: Second argument will not work.  Cannot hide progress bar.
    zipfolder.CopyHere(fso.GetAbsolutePathName(files[i]), 4 + 16);
    // FIXME: Wait until finish.  It seems that file size is updated
    // when all items in the folder are finished, instead of for each
    // item.
    for (j = 0; j < 10; ++j) {
      if (oldsize !== f.Size) {
        break;
      } else if (thread_count(process_id) > oldcount) {
        // CopyHere threads are running.  Wait it.
        break;
      }
      WScript.Sleep(1000);
    }
    // Wait until CopyHere threads are finished.
    while (thread_count(process_id) > oldcount) {
      WScript.Sleep(1000);
    }
    // If file size is not changed, compression is probably cancelled.
    if (oldsize === f.Size) {
      return false;
    }
  }

  return true;
}

function get_process_id() {
  var shell = WScript.CreateObject('WScript.Shell');
  var p = shell.Exec('cmd.exe');
  var wmi = GetObject('winmgmts://./root/cimv2');
  var e = new Enumerator(wmi.ExecQuery('SELECT * FROM Win32_Process WHERE ProcessId = ' + p.ProcessID));
  var process_id = null;
  for (; !e.atEnd(); e.moveNext()) {
    var item = e.item();
    process_id = item.ParentProcessId;
  }
  // Terminate() is slow.
  p.StdIn.Close();
  return process_id;
}

function thread_count(process_id) {
  var wmi = GetObject('winmgmts://./root/cimv2');
  var res = wmi.ExecQuery('SELECT * FROM Win32_Thread WHERE ProcessHandle = ' + process_id);
  return res.Count;
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
      recursive: true // placeholder.  Always recursive.
    }
  };
  var i;
  for (i = 0; i < args.length; ++i) {
    if (args[i].match(/^(-h|--help)$/)) {
      res.option.help = true;
    } else if (args[i].match(/^(-r)$/)) {
      res.option.recursive = true;
    } else if (args[i].match(/^-/)) {
      throw Error('unknown option: ' + args[i]);
    } else {
      res.args.push(args[i]);
    }
  }
  return res;
}

function usage() {
  WScript.Echo('usage: zip [options] zipfile file ...');
  WScript.Echo('  -h --help    print this text');
  WScript.Quit(1);
}

function main() {
  var args, zipfile, files;
  args = parse_arguments(getargs());
  if (args.option.help || args.args.length <= 1) {
    usage();
  }
  zipfile = args.args.shift();
  files = args.args;
  zip(zipfile, files);
}

main();
