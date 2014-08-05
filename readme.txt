
Delphi 6で「元に戻す」を実行するサンプルプログラムです。

まず、「ファイルを開く(File Open Dialog)」でファイルを開くのダイアログが開きます。
ここでファイルの操作を行って下さい。

「元に戻すを取得(Get Undo)」で元に戻す動作を表示し、
「元に戻す(Undo)」で元に戻します。


■簡単なUndoの実装方法の説明
http://www.eonet.ne.jp/~gakana/memo/xfmemo1.html


単純に元に戻すを行うだけであれば、IShellBrowserの実装はGetWindow以外はE_NOTIMPLでOKです。

function TForm1.GetWindow(out wnd: HWnd): HResult; stdcall;
begin
  wnd := Handle;
  Result := S_OK;
end;


Copyright (c) 2014 Gaku
Licensed under the MIT License:
http://www.opensource.org/licenses/mit-license.php


//Build
Borland Delphi 6 Personal
