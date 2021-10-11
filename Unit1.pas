////////////////////////////////////////////////////////////////////////////////
//
//  ****************************************************************************
//  * Unit Name : Unit1
//  * Purpose   : ���� ������ � ��������
//  * Author    : ��������� (Rouse_) ������
//  * Version   : 1.01
//  * Web Site  : http://rouse.front.ru
//  ****************************************************************************
//

unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ShellAPI, ShlObj, ActiveX, ComCtrls, CommCtrl, Menus, 
  ExtCtrls, CheckLst, ComObj, XPMan;

// ������� ���������� �� ��� ���������� �� ���������� ��������
// � ������ 6 �������.
// � ���������������� ���� ������� ������...
const
  DETAIL_COUNT = 11;
  WM_SHELLNOTIFIER = WM_USER;

const
  SHERB_NOCONFIRMATION  =  $0001;
  SHERB_NOPROGRESSUI    =  $0002;
  SHERB_NOSOUND         =  $0004;

  SHCNF_ACCEPT_INTERRUPTS     = $0001;
  SHCNF_ACCEPT_NON_INTERRUPTS = $0002;
  SHCNRF_RECURSIVEINTERRUPT   = $0004;

type
  TfrmRecycleBin = class(TForm)
    lvData: TListView;
    pmOperations: TPopupMenu;
    mnuRestore: TMenuItem;
    mnuSeparator1: TMenuItem;
    mnuSeparator2: TMenuItem;
    mnuDelete: TMenuItem;
    mnuPropertyes: TMenuItem;
    gbData: TGroupBox;
    lblElements: TLabel;
    lblSize: TLabel;
    btnEmpty: TButton;
    Bevel1: TBevel;
    edFileOrFolderPath: TEdit;
    btnDelete: TButton;
    Bevel2: TBevel;
    clbDrives: TCheckListBox;
    cbDellFromAllDrives: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnEmptyClick(Sender: TObject);
    procedure edFileOrFolderPathChange(Sender: TObject);
    procedure cbDellFromAllDrivesClick(Sender: TObject);
    procedure clbDrivesMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure clbDrivesKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormDestroy(Sender: TObject);
    procedure lvDataDblClick(Sender: TObject);
    procedure mnuRestoreClick(Sender: TObject);
  private
    HShellNotifyer: ULONG;
    ImageListHandle: THandle;
    procedure ViewRecycleBin;
    function ExecuteVerb(const VerbIndex: Byte): Boolean;
    procedure GetRecycleBinInfo;
    procedure FillDriveList;
    procedure UpdateEmptyButtonState;
    procedure SetRecycleBinNotifyer(const Logged: Boolean);
    procedure WMShellNotifyer(var Message: TMessage); message WM_SHELLNOTIFIER;
  end;

var
  frmRecycleBin: TfrmRecycleBin;

implementation

{$R *.dfm}

type
  TSHQueryRBInfo = packed record
    cbSize      : DWORD;
    i64Size,
    i64NumItems : TLargeInteger;
  end;
  PSHQueryRBInfo = ^TSHQueryRBInfo;

  TSHChangeNotifyEntry = packed record
    pidl: PItemIDList;
    fRecursive: BOOL;
  end;
  PSHChangeNotifyEntry = ^TSHChangeNotifyEntry;

  function SHEmptyRecycleBin(hwnd: HWND; pszRootPath: PChar;
    dwFlags: DWORD): HRESULT; stdcall;
    external 'shell32.dll' name 'SHEmptyRecycleBinA';

  function SHQueryRecycleBin(pszRootPath: PChar;
    var SHQueryRBInfo: TSHQueryRBInfo): HRESULT; stdcall;
    external  'Shell32.dll' name 'SHQueryRecycleBinA';

  function SHChangeNotifyRegister(hwnd: HWND; fSources: Byte; fEvents: LongInt;
    wMsg: UINT; cEntries: Byte; pfsne: PSHChangeNotifyEntry): ULONG; stdcall;
    external  'Shell32.dll';

  function SHChangeNotifyDeregister(uiID: ULONG): BOOL; stdcall;
    external  'Shell32.dll';

  procedure SHGetSetSettings(var lpss: TShellFlagState; dwMask: DWORD;
    bState: BOOL); stdcall; external  'Shell32.dll';

// ������� ����� �� QDialogs...
function StrRetToString(PIDL: PItemIDList; StrRet: TStrRet;
  Flag: String = ''): String;
var
  P: PChar;
begin
  case StrRet.uType of
    STRRET_CSTR:
      SetString(Result, StrRet.cStr, lStrLen(StrRet.cStr));
    STRRET_OFFSET:
      begin
        P := @PIDL.mkid.abID[StrRet.uOffset - SizeOf(PIDL.mkid.cb)];
        SetString(Result, P, PIDL.mkid.cb - StrRet.uOffset);
      end;
    STRRET_WSTR:
      if Assigned(StrRet.pOleStr) then
        Result := StrRet.pOleStr
      else
        Result := '';
  end;
  { This is a hack bug fix to get around Windows Shell Controls returning
    spurious "?"s in date/time detail fields }
  if (Length(Result) > 1) and (Result[1] = '?') and (Result[2] in ['0'..'9']) then
    Result := StringReplace(Result, '?', '', [rfReplaceAll]);
end;  

{ TfrmRecycleBin }

// �������� ����� ��� ����� � �������.
procedure TfrmRecycleBin.btnDeleteClick(Sender: TObject);
var
  Struct: TSHFileOpStruct;
begin
  with Struct do
  begin
    Wnd := Handle;
    wFunc := FO_DELETE;
    // Struct.pFrom - ������ ������������� ����� �������������� ������!
    pFrom := PChar(edFileOrFolderPath.Text + #0);
    pTo := nil;
    fFlags := FOF_ALLOWUNDO;
    fAnyOperationsAborted := True;
    hNameMappings := nil;
    lpszProgressTitle := nil;
  end;
  OleCheck(SHFileOperation(Struct));
end;

// ������� �������
procedure TfrmRecycleBin.btnEmptyClick(Sender: TObject);
var
  Err: HRESULT;
  I: Integer;
begin
  Err := S_FALSE;
  if not cbDellFromAllDrives.Checked then
  begin
    // ������� ������ ��������� ������
    for I := 0 to clbDrives.Items.Count - 1 do
      if clbDrives.Checked[I] then
        if not (Err in [S_OK, S_FALSE]) then
          RaiseLastOSError
        else
          if Err = S_FALSE then
            Err := SHEmptyRecycleBin(Handle,
              PChar(clbDrives.Items.Strings[I]), SHERB_NOSOUND)
          else
            Err := SHEmptyRecycleBin(Handle,
              PChar(clbDrives.Items.Strings[I]), SHERB_NOCONFIRMATION or SHERB_NOSOUND);
  end
  else
    // ������� ���� ������
    Err := SHEmptyRecycleBin(Handle, nil, SHERB_NOSOUND);
  OleCheck(Err);
end;

procedure TfrmRecycleBin.cbDellFromAllDrivesClick(Sender: TObject);
begin
  clbDrives.Enabled := not cbDellFromAllDrives.Checked;
  UpdateEmptyButtonState;
end;

procedure TfrmRecycleBin.clbDrivesKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  clbDrives.Perform(WM_LBUTTONUP, 0, 0);
end;

procedure TfrmRecycleBin.clbDrivesMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  Application.ProcessMessages;
  UpdateEmptyButtonState;
end;

procedure TfrmRecycleBin.edFileOrFolderPathChange(Sender: TObject);
begin
  // ���������� ������ ������ � ������ ��������� ����
  btnDelete.Enabled := FileExists(edFileOrFolderPath.Text) or
    DirectoryExists(edFileOrFolderPath.Text);
end;

// ������� �������� �������� ��� ��������� � ListView ���������
function TfrmRecycleBin.ExecuteVerb(const VerbIndex: Byte): Boolean;

  function GetLVItemText(const ItemIndex, SectionIndex: Integer): String;
  begin
    if SectionIndex = 0 then
      Result := lvData.Items.Item[ItemIndex].Caption
    else
      Result := lvData.Items.Item[ItemIndex].SubItems.Strings[SectionIndex - 1];
  end;

const
  VerbData: array [0..2] of String = ('undelete', 'delete', 'properties');

var
  ppidl, Item: PItemIDList;
  ResultItems: array of PItemIDList;
  Desktop: IShellFolder;
  RecycleBin: IShellFolder2;
  RecycleBinEnum: IEnumIDList;
  Fetched, I, Z, PIDLCount: Cardinal;
  Details: TShellDetails;
  Mallok: IMalloc;
  Valid: Boolean;
  Context: IContextMenu;
  AInvokeCommand: TCMInvokeCommandInfo;
begin
  Result := False;
  ResultItems := nil;
  PIDLCount := 0;
  // �������� ��������� ��� ������ �������� ����� ����������� ������� ������
  OleCheck(SHGetMalloc(Mallok));
  // �������� ��������� �� �������
  OleCheck(SHGetSpecialFolderLocation(Handle,
    CSIDL_BITBUCKET, ppidl));
  // �������� ��������� �� ������� ����
  OleCheck(SHGetDesktopFolder(Desktop));
  // �������� ��������� �� �������
  OleCheck(Desktop.BindToObject(ppidl, nil,
    IID_IShellFolder2, RecycleBin));
  // �������� ��������� ��� ������������ ��������� �������
  OleCheck(RecycleBin.EnumObjects(Handle,
    SHCONTF_FOLDERS or SHCONTF_NONFOLDERS or SHCONTF_INCLUDEHIDDEN,
    RecycleBinEnum));
  // ���������� ���������� �������
  for Z := 0 to lvData.Items.Count - 1 do
  begin
    RecycleBinEnum.Next(1, Item, Fetched);
    if Fetched = 0 then Break;
    Valid := False;
    // ����������� ������ ���������� ��������
    if lvData.Items.Item[Z].Selected then
      for I := 0 to DETAIL_COUNT - 1 do
        if RecycleBin.GetDetailsOf(Item, I, Details) = S_OK then
        try
          // ���� ������ ��� �������
          Valid := GetLVItemText(Z, I) = StrRetToString(Item, Details.str);
          if not Valid then Break;
        finally
          Mallok.Free(Details.str.pOleStr);
        end;
    if Valid then
    begin
      SetLength(ResultItems, Length(ResultItems) + 1);
      ResultItems[Length(ResultItems) - 1] := Item;
      Inc(PIDLCount);
    end;
  end;
  // ���� ���������� ������� ������
  if ResultItems <> nil then         
  begin
    // ���������� � ��� �������� ��� ������ ���������� IContextMenu
    if RecycleBin.GetUIObjectOf(Handle, PIDLCount, ResultItems[0],
      IID_IContextMenu, nil, Pointer(Context)) = S_OK then
    begin
      FillMemory(@AInvokeCommand, SizeOf(AInvokeCommand), 0);
      with AInvokeCommand do
      begin
        cbSize := SizeOf(AInvokeCommand);
        hwnd := Handle;
        lpVerb := PChar(VerbData[VerbIndex]); // ��������� ��������� ��� �������� ��� ���������...
        fMask := CMIC_MASK_FLAG_NO_UI;
        nShow := SW_SHOWNORMAL;
      end;
      // ���������� �������...
      Result := Context.InvokeCommand(AInvokeCommand) = S_OK;
    end;
  end;
end;

// ��������� ������ �������������� ���������� ������
procedure TfrmRecycleBin.FillDriveList;
const
  NameSize = 4;
  VolumeCount = 26;
  TotalSize = NameSize * VolumeCount;
var
  Buff, Volume: String;
  I, Count: Integer;
begin
  SetLength(Buff, TotalSize);
  Count := GetLogicalDriveStrings(TotalSize, @Buff[1]) div NameSize;
  if Count > 0 then
    for I := 0 to Count - 1 do
    begin
      Volume := PChar(@Buff[(I * NameSize) + 1]);
      if  GetDriveType(PChar(Volume)) = DRIVE_FIXED then
        clbDrives.Items.Add(Volume);
    end;
end;

procedure TfrmRecycleBin.FormCreate(Sender: TObject);
var
  FileInfo: TSHFileInfo;
begin
  lvData.DoubleBuffered := True;
  // ��������� ���������� � ������� ��������� �������.
  FillDriveList;
  SetRecycleBinNotifyer(True);
  ViewRecycleBin;
  ImageListHandle := SHGetFileInfo('C:\', 0, FileInfo, SizeOf(FileInfo),
    SHGFI_SYSICONINDEX or SHGFI_SMALLICON);
  SendMessage(lvData.Handle, LVM_SETIMAGELIST, LVSIL_SMALL, ImageListHandle);
end;

procedure TfrmRecycleBin.FormDestroy(Sender: TObject);
begin
  SetRecycleBinNotifyer(False);
  ImageList_Destroy(ImageListHandle);
end;

// ��������� ���������� � ������� ����� SHQueryRecycleBin
procedure TfrmRecycleBin.GetRecycleBinInfo;
var
  Info: TSHQueryRBInfo;
  Err: HRESULT;
begin
  ZeroMemory(@Info, SizeOf(Info));
  Info.cbSize := SizeOf(Info);
  // ������ ���������� �������� ������ ������ ������ ��� ��������� ������
  // � �������� ���� ��������� ������.
  // ���� �������� ������ ���������� nil, �� ��� Windows� 2000 ������ �������
  // ������ E_INVALIDARG, � ��� ��� �� �����...
  Err := SHQueryRecycleBin('', Info);
  if Err = S_OK then
  begin
    lblElements.Caption := Format('Elements count: %d', [Info.i64NumItems]);
    lblSize.Caption := Format('Total sise: %d Mb', [Info.i64Size div 1048576]);
  end;
end;

procedure TfrmRecycleBin.lvDataDblClick(Sender: TObject);
begin
  mnuPropertyes.Click;
end;

procedure TfrmRecycleBin.mnuRestoreClick(Sender: TObject);
begin
  // ����� �������� ��� ���� ��������� ����...
  // (������������, �������, ��������)
  if lvData.Selected <> nil then
    if not ExecuteVerb(TMenuItem(Sender).Tag) then RaiseLastOSError;
end;

// ������������� �������� �� �������� � ��������� ����������� WM_SHELLNOTIFIER,
// �� ������ �������� ����� ������������ �������� �������...
procedure TfrmRecycleBin.SetRecycleBinNotifyer(const Logged: Boolean);
var
  pidl: PItemIDList;
  Notifier: TSHChangeNotifyEntry;
begin
  OleCheck(SHGetSpecialFolderLocation(Handle, CSIDL_BITBUCKET, pidl));
  Notifier.fRecursive := True;
  Notifier.pidl := pidl;
  if Logged then
  begin
    HShellNotifyer := SHChangeNotifyRegister(Handle, SHCNF_ACCEPT_INTERRUPTS or
    SHCNF_ACCEPT_NON_INTERRUPTS or SHCNRF_RecursiveInterrupt, SHCNE_ALLEVENTS,
    WM_SHELLNOTIFIER, 1, @Notifier);
    if HShellNotifyer = 0 then RaiseLastOSError;
  end
  else
    if not SHChangeNotifyDeregister(HShellNotifyer) then
      RaiseLastOSError;
end;

procedure TfrmRecycleBin.UpdateEmptyButtonState;
var
  I: Integer;
  IsEnable: Boolean;
begin
  IsEnable := False;
  if cbDellFromAllDrives.Checked then
    IsEnable := True
  else
    for I := 0 to clbDrives.Items.Count - 1 do
      if clbDrives.Checked[I] then
      begin
        IsEnable := True;
        Break;
      end;
  btnEmpty.Enabled := IsEnable;
end;

// ������� ���������� �������...
procedure TfrmRecycleBin.ViewRecycleBin;
var
  ppidl, Item: PItemIDList;
  Desktop: IShellFolder;
  RecycleBin: IShellFolder2;
  RecycleBinEnum: IEnumIDList;
  Fetched, I: Cardinal;
  Details: TShellDetails;
  Mallok: IMalloc;
  TmpStr: ShortString;
  FileInfo: TSHFileInfo;
begin
  GetRecycleBinInfo;
  lvData.Items.BeginUpdate;
  try
    // ������������� ��������� ListView
    lvData.Clear;
    lvData.Columns.Clear;
    lvData.ViewStyle := vsReport;
    // �������� ��������� ��� ������ �������� ����� ����������� ������� ������
    OleCheck(SHGetMalloc(Mallok));
    // �������� ��������� �� �������
    OleCheck(SHGetSpecialFolderLocation(Handle,
      CSIDL_BITBUCKET, ppidl));
    // �������� ��������� �� ������� ����
    OleCheck(SHGetDesktopFolder(Desktop));
    // �������� ��������� �� �������
    OleCheck(Desktop.BindToObject(ppidl, nil,
      IID_IShellFolder2, RecycleBin));
    // �������� ��������� ��� ������������ ��������� �������
    OleCheck(RecycleBin.EnumObjects(Handle,
      SHCONTF_FOLDERS or SHCONTF_NONFOLDERS or SHCONTF_INCLUDEHIDDEN,
      RecycleBinEnum));
    // ������� �������
    for I := 0 to DETAIL_COUNT - 1 do
      if RecycleBin.GetDetailsOf(nil, I, Details) = S_OK then
      try
        with lvData.Columns.Add do
        begin
          Caption := StrRetToString(Item, Details.str);
          Width := lvData.Canvas.TextWidth(Caption) + 24;
        end;
      finally
        Mallok.Free(Details.str.pOleStr);
      end;

    // ���������� ���������� �������
    while True do
    begin
      //����� ������ ���� ��������� ������� �������
      RecycleBinEnum.Next(1, Item, Fetched);
      if Fetched = 0 then Break;
      // �������� ���������� � ��������
      if RecycleBin.GetDetailsOf(Item, 0, Details) = S_OK then
      begin
        try
          // �������� ��� ��������
          TmpStr := StrRetToString(Item, Details.str);
          // �������� ������ ������ �������� � ��������� �����
          SHGetFileInfo(PChar(Item), 0, FileInfo, SizeOf(FileInfo),
            SHGFI_PIDL or SHGFI_SYSICONINDEX);
        finally
          // ����������� ������
          Mallok.Free(Details.str.pOleStr);
        end;

        // ��������� ������� � ��� ��������� � ������
        with lvData.Items.Add do
        begin
          Caption := TmpStr;
          ImageIndex := FileInfo.iIcon;
          for I := 1 to DETAIL_COUNT - 1 do
            if RecycleBin.GetDetailsOf(Item, I, Details) = S_OK then
            try
              SubItems.Add(StrRetToString(Item, Details.str));
            finally
              Mallok.Free(Details.str.pOleStr);
            end;
        end;
      end;
    end;
  finally
    lvData.Items.EndUpdate;
  end;
end;

procedure TfrmRecycleBin.WMShellNotifyer(var Message: TMessage);
begin
  // ������ ����������� � ��������� ������� - ������������� �� ��������
  ViewRecycleBin;
end;

end.

