unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls,
//追加
  ShlObj, ActiveX, CommCtrl, StrUtils;

const
  CommandID_UNDO = 28700;
  WM_GETISHELLBROWSER = WM_USER + 7;
  FWF_NOWEBVIEW = $00010000;

type
  TForm1 = class(TForm, IShellBrowser)
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    OpenDialog1: TOpenDialog;
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure Button2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
  private
    { Private 宣言 }
    FShellView: IShellView;
  public
    { Public 宣言 }
    procedure DoUndo;
    function GetUndo: string;
    function CreateViewWindowEx(out ShellView: IShellView): Boolean;
    procedure DestroyViewWindowEx;
  protected
    { Protected 宣言 }
{ IShellBrowser }
    procedure WMGetIShellBrowser(var msg: TMessage);
      message WM_GETISHELLBROWSER;
    function InsertMenusSB(hMenuShared: HMENU;
      out MenuWidths: TOleMenuGroupWidths): HResult; overload; stdcall;
    function SetMenuSB(hMenuShared: HMENU;
      hOleMenuReserved: HOLEMENU; hwndActiveObject: HWND): HResult; stdcall;
    function RemoveMenusSB(hMenuShared: HMENU): HResult; stdcall;
    function SetStatusTextSB(StatusText: POleStr): HResult; stdcall;
    function EnableModelessSB(Enable: BOOL): HResult; stdcall;
    function TranslateAcceleratorSB(Msg: PMsg; ID: Word): HResult; stdcall;
    function BrowseObject(pidl: PItemIDList; flags: UINT): HResult; stdcall;
    function GetViewStateStream(Mode: DWORD; out Stream: IStream): HResult; stdcall;
    function GetControlWindow(ID: UINT; out Wnd: HWND): HResult; stdcall;
    function SendControlMsg(ID, Msg: UINT; wParm: WPARAM; lParm: LPARAM;
      var Rslt: LResult): HResult; stdcall;
    function QueryActiveShellView(var ShellView: IShellView): HResult; stdcall;
    function OnViewWindowActive(var ShellView: IShellView): HResult; stdcall;
    function SetToolbarItems(TBButton: PTBButton;
      nButtons, uFlags: UINT): HResult; stdcall;
    function GetWindow(out wnd: HWnd): HResult; stdcall;
    function ContextSensitiveHelp(fEnterMode: BOOL): HResult; stdcall;

  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

//WM_GETISHELLBROWSERを処理しないとWindows 95/98/NTでエラー
procedure TForm1.WMGetIShellBrowser(var msg: TMessage);
begin
  msg.Result := Integer(Self as IShellBrowser);
end;

{ IShellBrowser }

function TForm1.InsertMenusSB(hMenuShared: HMENU;
  out MenuWidths: TOleMenuGroupWidths): HResult; stdcall;
begin
  Result := E_NOTIMPL;
end;

function TForm1.SetMenuSB(hMenuShared: HMENU;
  hOleMenuReserved: HOLEMENU; hwndActiveObject: HWND): HResult; stdcall;
begin
  Result := E_NOTIMPL;
end;

function TForm1.RemoveMenusSB(hMenuShared: HMENU): HResult; stdcall;
begin
  Result := E_NOTIMPL;
end;

function TForm1.SetStatusTextSB(StatusText: POleStr): HResult; stdcall;
begin
  Result := E_NOTIMPL;
end;

function TForm1.EnableModelessSB(Enable: BOOL): HResult; stdcall;
begin
  Result := E_NOTIMPL;
end;

function TForm1.TranslateAcceleratorSB(Msg: PMsg; ID: Word): HResult; stdcall;
begin
  Result := E_NOTIMPL;
end;

function TForm1.BrowseObject(pidl: PItemIDList; flags: UINT): HResult; stdcall;
begin
  Result := E_NOTIMPL;
end;

function TForm1.GetViewStateStream(Mode: DWORD; out Stream: IStream): HResult; stdcall;
begin
  Result := E_NOTIMPL;
end;

function TForm1.GetControlWindow(ID: UINT; out Wnd: HWND): HResult; stdcall;
begin
  Result := E_NOTIMPL;
end;

function TForm1.SendControlMsg(ID, Msg: UINT; wParm: WPARAM; lParm: LPARAM;
  var Rslt: LResult): HResult; stdcall;
begin
  Result := E_NOTIMPL;
end;

function TForm1.QueryActiveShellView(var ShellView: IShellView): HResult; stdcall;
begin
  Result := E_NOTIMPL;
  if Assigned(FShellView) then
  begin
    Pointer(ShellView) := nil;  //参照を消去
    ShellView := FShellView;
    Result := S_OK;
  end;
end;

function TForm1.OnViewWindowActive(var ShellView: IShellView): HResult; stdcall;
begin
  Result := E_NOTIMPL;
end;

function TForm1.SetToolbarItems(TBButton: PTBButton;
  nButtons, uFlags: UINT): HResult; stdcall;
begin
  Result := E_NOTIMPL;
end;

function TForm1.GetWindow(out wnd: HWnd): HResult; stdcall;
begin
  wnd := Handle;
  Result := S_OK;
end;

function TForm1.ContextSensitiveHelp(fEnterMode: BOOL): HResult; stdcall;
begin
  Result := E_NOTIMPL;
end;

//元に戻すを実行
procedure TForm1.DoUndo;
var
  hExplorer: HWND;
  ShellView: IShellView;
begin
  if CreateViewWindowEx(ShellView) then
  begin
    ShellView.UIActivate(Cardinal(SVUIA_ACTIVATE_NOFOCUS));
    ShellView.GetWindow(hExplorer);
    SendMessage(hExplorer, WM_COMMAND, CommandID_UNDO - 1, 0);
  end;
end;

//元に戻すを取得 (Windows 98以降)
function TForm1.GetUndo: string;
var
  ShellView: IShellView;
  ContextMenu: IContextMenu;
  Menu: HMENU;
  szMenu: Array[0..63] of Char;
  i, nCount, nPos: Integer;
begin
  Result := '';
  if CreateViewWindowEx(ShellView) then
  begin
    //バックグラウンドのコンテキストメニューを取得
    if Succeeded(ShellView.GetItemObject(SVGIO_BACKGROUND, IID_IContextMenu, Pointer(ContextMenu))) then
    begin
      Menu := CreatePopupMenu;
      try
        if Succeeded(ContextMenu.QueryContextMenu(Menu, 0, 1, $7FFF, CMF_DEFAULTONLY)) then
        begin
          nCount := GetMenuItemCount(Menu); // Windows 95では取得できない
          for i := 0 to nCount - 1 do
          begin
            GetMenuString(Menu, i, szMenu, 64, MF_BYPOSITION);
            nPos := Pos('Ctrl+Z', szMenu);
            if (nPos > 0) or AnsiContainsText(szMenu, '元に戻す') or AnsiContainsText(szMenu, 'Undo') then
            begin
              if nPos > 0 then
              begin
                Result := TrimRight(Copy(szMenu, 1, nPos - 2));
              end
              else
              begin
                Result := szMenu;
              end;
              break;
            end;
          end;
        end;
      finally
        DestroyMenu(Menu);
      end;
    end;
  end;
end;

//ShellViewのViewWindowを作成
function TForm1.CreateViewWindowEx(out ShellView: IShellView): Boolean;
var
  fs: FOLDERSETTINGS;
  R: TRect;
  ShellFolder: IShellFolder;
  hExplorer: HWND;
begin
  if FShellView = nil then
  begin
    if Succeeded(SHGetDesktopFolder(ShellFolder)) then
    begin
      if Succeeded(ShellFolder.CreateViewObject(Handle, IID_IShellView, Pointer(FShellView))) then
      begin
        R := Rect(-2, -2, -1, -1);
        fs.ViewMode := FVM_LIST;
        fs.fFlags := FWF_NOWEBVIEW;
        if Failed(FShellView.CreateViewWindow(nil, fs, Self as IShellBrowser, R, hExplorer)) then
        begin
          FShellView := nil;
        end;
      end;
    end;
  end;
  ShellView := FShellView;
  Result := Assigned(ShellView);
end;

//ShellViewのViewWindowを閉じる
procedure TForm1.DestroyViewWindowEx;
var
  hExplorer: HWND;
begin
  if Assigned(FShellView) then
  begin
    try
      FShellView.UIActivate(Cardinal(SVUIA_DEACTIVATE));
    except
    end;
    if (Win32MajorVersion * 100 + Win32MinorVersion) >= 501 then
    begin
      //Windows MEだとデスクトップがちらつく
      FShellView.DestroyViewWindow;
    end;
    FShellView.GetWindow(hExplorer);
    if hExplorer <> 0 then
    begin
      PostMessage(hExplorer, WM_CLOSE, 0, 0);
    end;
  end;
  FShellView := nil;
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  DoUndo;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  ShowMessage(GetUndo);
end;

//終了時の処理
procedure TForm1.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  DestroyViewWindowEx;
end;

procedure TForm1.Button3Click(Sender: TObject);
begin
  OpenDialog1.Execute;
end;

end.
