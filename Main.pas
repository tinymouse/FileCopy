unit Main;

interface

uses
  Windows, SysUtils, Classes, Controls, Dialogs,
  SearchFile, Progress, XPMan;

type
  TIntegerList = class(TList)
  protected
    function GetData(Index: Integer): Integer;
    procedure SetData(Index: Integer; const Value: Integer);
  public
    destructor Destroy; override;
    procedure Clear;
    function Add(Value: Integer): Integer;
    property Items[Index: Integer]: Integer read GetData write SetData;
  end;

  TDateTimeList = class(TList)
  protected
    function GetData(Index: Integer): TDateTime;
    procedure SetData(Index: Integer; const Value: TDateTime);
  public
    destructor Destroy; override;
    procedure Clear;
    function Add(Value: TDateTime): Integer;
    property Items[Index: Integer]: TDateTime read GetData write SetData;
  end;

  TCommand = (cmdCopy, cmdMove, cmdAdd, cmdSync);

  TMainModule = class(TDataModule)
    XPManifest: TXPManifest;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private 宣言 }
    FSelectedFiles: TStringList;  { 選択されたファイル/ディレクトリのパス }
    FSelectedDirs: Boolean;  { 選択されたディレクトリがあるか }
    FSourceDir: String;
    FTargetDir: String;  { 複写/移動/追加/更新先のディレクトリのパス }
    FFilesToCopy: TStringList;  { 複写/移動の対象となるファイル/ディレクトリのパス }
    FFilesToDest: TStringList;
    FFilesToDelete: TStringList;  { 削除の対象となるファイル/ディレクトリのパス }
    FCreateAgesToCopy: TDateTimeList;  { 日時複写の対象となるファイル/ディレクトリの作成日時 }
    FLastAccessAgesToCopy: TDateTimeList;    { 日時複写の対象となるファイル/ディレクトリのアクセス日時 }
    FLastWriteAgesToCopy: TDateTimeList;  { 日時複写の対象となるファイル/ディレクトリの更新日時 }
    FFilesToCopyAge: TStringList;  { 日時複写の対象となるファイル/ディレクトリの複写/移動先のパス }
    FAttrsToCopy: TIntegerList;  { 属性複写の対象となるディレクトリの属性 }
    FDirsToCopyAttr: TStringList;  { 属性複写の対象となるディレクトリの複写/移動先のパス }
    FMove: Boolean;  { 移動か否か }
    FCopyAge: Boolean;  { 日時/属性を複写するか }
    FConfirm: Boolean;  { 上書確認等するか }
    FForce: Boolean;  { 強制的にフォルダの日時を変更するか }
    FCopyCreateAge: Boolean;  { 作成日時は保持するか }
    FCopyNewerCreateAge: Boolean;  { コピー先の作成日時がコピー元より古いとき日時をコピーするか }
    FCopyLastAccessAge: Boolean;  { アクセス日時も保持するか }
    FCopyDirAge: Boolean;  { ディレクトリの日時も保持するか }
    FSearchFile: TSearchFile;
    FProgressDlg: TProgressDlg;
    procedure DoCopy;
    procedure DoMove;
    procedure FindFilesToCopy;
    procedure CheckFileToCopy(Name: String);
    procedure FoundFileToCopy(Sender: TObject; Name: String;
      Data: TSearchFileData; var Continue: Boolean);
    procedure OperateFilesToCopy;
    procedure CopyAgesToCopy;
    procedure CopyAttrsToCopy;
    procedure DoAdd;
    procedure FindFilesToAdd;
    procedure CheckFileToAdd(Name: String);
    procedure FoundFileToAdd(Sender: TObject; Name: String;
      Data: TSearchFileData; var Continue: Boolean);
    procedure DoUpdate;
    procedure FindFilesToDelete;
    procedure FoundFileToDelete(Sender: TObject; Name: String;
      Data: TSearchFileData; var Continue: Boolean);
    procedure OperateFilesToDelete;
    procedure DoBackup;
    procedure DoCopyAge;
  public
    { Public 宣言 }
    procedure RunCommand;
  end;

var
  MainModule: TMainModule;

implementation

uses
  Forms, FileCtrl, ShellApi, StrUtils, Utils;

{$R *.DFM}

{ TIntegerList }

destructor TIntegerList.Destroy;
begin
  Clear;
  inherited Destroy;
end;

procedure TIntegerList.Clear;
var
  n: Integer;
begin
  for n := 0 to Count - 1 do
    Dispose(PInteger(inherited Items[n]));
  inherited Clear;
end;

function TIntegerList.Add(Value: Integer): Integer;
var
  NewData: PInteger;
begin
  New(NewData);
  NewData^ := Value;
  Result := inherited Add(NewData);
end;

function TIntegerList.GetData(Index: Integer): Integer;
begin
  Result := Integer((inherited Items[Index])^);
end;

procedure TIntegerList.SetData(Index: Integer; const Value: Integer);
begin
  Integer((inherited Items[Index])^) := Value;
end;

{ TDateTimeList }

destructor TDateTimeList.Destroy;
begin
  Clear;
  inherited Destroy;
end;

procedure TDateTimeList.Clear;
var
  n: Integer;
begin
  for n := 0 to Count - 1 do
    Dispose(PDateTime(inherited Items[n]));
  inherited Clear;
end;

function TDateTimeList.Add(Value: TDateTime): Integer;
var
  NewData: PDateTime;
begin
  New(NewData);
  NewData^ := Value;
  Result := inherited Add(NewData);
end;

function TDateTimeList.GetData(Index: Integer): TDateTime;
begin
  Result := TDateTime((inherited Items[Index])^);
end;

procedure TDateTimeList.SetData(Index: Integer; const Value: TDateTime);
begin
  TDateTime((inherited Items[Index])^) := Value;
end;

{ TMainModule }

procedure TMainModule.DataModuleCreate(Sender: TObject);
begin
  { 変数を初期化 }
  FSelectedFiles := TStringList.Create;
  FFilesToCopy := TStringList.Create;
  FFilesToDest := TStringList.Create;
  FFilesToDelete := TStringList.Create;
  FCreateAgesToCopy := TDateTimeList.Create;
  FLastAccessAgesToCopy := TDateTimeList.Create;
  FLastWriteAgesToCopy := TDateTimeList.Create;
  FFilesToCopyAge := TStringList.Create;
  FAttrsToCopy := TIntegerList.Create;
  FDirsToCopyAttr := TStringList.Create;
  FProgressDlg := TProgressDlg.Create(Self);
  FSearchFile := TSearchFile.Create(Self)
end;

procedure TMainModule.DataModuleDestroy(Sender: TObject);
begin
  { 変数を解放 }
  FSelectedFiles.Free;
  FFilesToCopy.Free;
  FFilesToDest.Free;
  FFilesToDelete.Free;
  FCreateAgesToCopy.Free;
  FLastAccessAgesToCopy.Free;
  FLastWriteAgesToCopy.Free;
  FFilesToCopyAge.Free;
  FAttrsToCopy.Free;
  FDirsToCopyAttr.Free;
  FProgressDlg.Free;
  FSearchFile.Free;
end;

procedure TMainModule.RunCommand;
var
  buflist: TStringList;
  n: Integer;
begin
  if ParamCount < 1 then
  begin
    MessageDlg('コマンド指定が間違っています', mtError, [mbOK], 0);
    Exit;
  end;

  { 複写/移動元と複写/移動先をコマンドライン変数から取得 }
  buflist := TStringList.Create;
  for n := 2 to ParamCount do
    if ParamStr(n)[1] <> '-' then
      buflist.Add(ParamStr(n));
  FSelectedFiles.Clear;
  FSelectedDirs := False;
  for n := 0 to buflist.Count - 2 do
  begin
    FSelectedFiles.Add(buflist.Strings[n]);
    if DirectoryExists(buflist.Strings[n]) then
      FSelectedDirs := True;
  end;
  FTargetDir := buflist.Strings[buflist.Count - 1];
  buflist.Free;
  { 取得されたファイル/ディレクトリの数が 0 なら }
  if FSelectedFiles.Count = 0 then
  begin
    MessageDlg('複写/移動するファイル/ディレクトリがありません', mtError, [mbOK], 0);
    Exit;
  end;

  { オプション指定をコマンドライン変数から取得 }
  FCopyAge := FindCmdLineSwitch('copyage');
  FConfirm := not FindCmdLineSwitch('noconfirm');
  FForce := FindCmdLineSwitch('force');
  FCopyCreateAge := not FindCmdLineSwitch('notcopycreateage');
  FCopyNewerCreateAge := not FindCmdLineSwitch('notcopynewercreateage');
  FCopyLastAccessAge := FindCmdLineSwitch('copylastaccessage');
  FCopyDirAge := not FindCmdLineSwitch('notcopydirage');

  { ディレクトリが選択されていて、}
  { 日時を複写するようになっていて、}
  { 複写先がネットワーク上のディレクトリなら }
  if FSelectedDirs
    and FCopyAge
    and PathIsUNC(PChar(ExpandUNCFileName(FTargetDir))) then
  begin
    { 強制的にフォルダの日時を変更する指定があり、NT 系 Windows なら }
    if FForce and OsIsWinNt then
      { 何もしない }
    else
      with CreateMessageDialog(QuotedStr(FTargetDir) + ' はネットワーク上のフォルダなので、'
        + 'フォルダの日時は保持することができません'
        + #13#10#13#10 + '処理を続けますか', mtError, [mbYes, mbNo]) do
      try
        ParentWindow := GetDesktopWindow;
        if ShowModal = mrNo then
          Exit;
      finally
        Free;
      end;
  end;

  { 指令をコマンドライン変数から取得 }
  if FindCmdLineSwitch('copy', True) then
    DoCopy
  else if FindCmdLineSwitch('move', True) then
    DoMove
  else if FindCmdLineSwitch('add', True) then
    DoAdd
  else if FindCmdLineSwitch('update', True) then
    DoUpdate
  else if FindCmdLineSwitch('backup', True) then
    DoBackup
  else if FindCmdLineSwitch('copyage', True) then
    DoCopyAge
  else
    MessageDlg('コマンド指定が間違っています', mtError, [mbOK], 0);
end;

procedure TMainModule.DoCopy;
begin
  try
    { 移動しない指定 }
    FMove := False;
    { 複写の対象になるファイル/ディレクトリを検索 }
    FindFilesToCopy;
    { 複写の対象のファイル/ディレクトリを複写 }
    OperateFilesToCopy;
    { 日時複写の対象のファイル/ディレクトリの日時を複写 }
    CopyAgesToCopy;
    { 属性複写の対象のファイル/ディレクトリの属性を複写 }
    CopyAttrsToCopy;
  except
    MessageDlg('ファイル/フォルダのコピーを中断しました', mtError, [mbOK], 0);
  end;
end;

procedure TMainModule.DoMove;
begin
  try
    { 移動する指定 }
    FMove := True;
    { 移動の対象になるファイル/ディレクトリを検索 }
    FindFilesToCopy;
    { 移動の対象のファイル/ディレクトリを移動 }
    OperateFilesToCopy;
    { 日時複写の対象のファイル/ディレクトリの日時を複写 }
    CopyAgesToCopy;
    { 属性複写の対象のファイル/ディレクトリの属性を複写 }
    CopyAttrsToCopy;
  except
    MessageDlg('ファイル/フォルダの移動を中断しました', mtError, [mbOK], 0);
  end;
end;

procedure TMainModule.FindFilesToCopy;
var
  n: Integer;
begin
  { 経過ダイアログを表示 }
  FProgressDlg.Show;
  if FMove then
    FProgressDlg.Caption := '移動するﾌｧｲﾙ/ﾌｫﾙﾀﾞの検索'
  else
    FProgressDlg.Caption := '複写するﾌｧｲﾙ/ﾌｫﾙﾀﾞの検索';
  FProgressDlg.InitCount;
  { 変数を初期化 }
  FFilesToCopy.Clear;
  FFilesToDest.Clear;
  FCreateAgesToCopy.Clear;
  FLastWriteAgesToCopy.Clear;
  FFilesToCopyAge.Clear;
  { 複写の対象になるファイル/ディレクトリを検索 }
  for n := 0 to FSelectedFiles.Count - 1 do
  begin
    FSourceDir := ExtractFilePath(FSelectedFiles.Strings[n]);
    CheckFileToCopy(FSelectedFiles.Strings[n]);
    { 下位のディレクトリを検索する }
    if not DirectoryExists(FSelectedFiles.Strings[n]) then
      Continue;
    FSourceDir := FSelectedFiles.Strings[n];
    with FSearchFile do
    begin
      Directory := FSourceDir;
      SearchName := '*.*';
      Recursive := True;
      OnFind := FoundFileToCopy;
      if not Search then Abort;
    end;
  end;
  { 経過ダイアログを破棄 }
  FProgressDlg.Hide;
end;

procedure TMainModule.CheckFileToCopy(Name: String);
var
  Dest: String;
  CreateAge, LastAccessAge, LastWriteAge: TDateTime;
begin
  Dest := IncludeTrailingPathDelimiter(FTargetDir)
    + ExtractRelativePath(IncludeTrailingPathDelimiter(FSourceDir), Name);
  { 経過ダイアログへ表示 }
  FProgressDlg.SetFileName(Name);
  FProgressDlg.AddCount;
  { 複写対象に追加 }
  FFilesToCopy.Add(Name);
  FFilesToDest.Add(Dest);
  { 日時を複写するようになっていれば }
  if FCopyAge then
  begin
    FileGetAge(Name, CreateAge, LastAccessAge, LastWriteAge);
    FCreateAgesToCopy.Add(CreateAge);
    FLastAccessAgesToCopy.Add(LastAccessAge);
    FLastWriteAgesToCopy.Add(LastWriteAge);
    FFilesToCopyAge.Add(Dest);
  end;
  { 日時を複写するようになっていて }
  { 複写先のディレクトリに複写対象と同じディレクトリがあり }
  { 複写先のディレクトリの属性が複写対象と違えば }
  if FCopyAge
    and DirectoryExists(Dest)
    and (FileGetAttr(Dest) <> FileGetAttr(Name)) then
  begin
    FAttrsToCopy.Add(FileGetAttr(Name));
    FDirsToCopyAttr.Add(Dest);
  end;
end;

procedure TMainModule.FoundFileToCopy(Sender: TObject; Name: String;
  Data: TSearchFileData; var Continue: Boolean);
var
  Dest: String;
  n: Integer;
  CreateAge, LastWriteAge, LastAccessAge: TDateTime;
begin
  Application.ProcessMessages;
  if FProgressDlg.Stop then Continue := False;

  Dest := IncludeTrailingPathDelimiter(FTargetDir)
    + IncludeTrailingPathDelimiter(ExtractFileName(ExcludeTrailingPathDelimiter(FSourceDir)))
    + ExtractRelativePath(IncludeTrailingPathDelimiter(FSourceDir), Name);
  { 経過ダイアログへ表示 }
  FProgressDlg.SetFileName(Name);
  FProgressDlg.AddCount;
  { 複写対象に追加 }
  for n := 0 to FFilesToCopy.Count - 1 do
    if Pos(FFilesToCopy.Strings[n], Name) = 1 then
      break;
  if n > FFilesToCopy.Count - 1 then  { 上位のディレクトリが複写対象になっていなければ }
  begin
    FFilesToCopy.Add(Name);
    FFilesToDest.Add(Dest);
  end;
  { 日時を複写するようになっていれば }
  if FCopyAge then
  begin
    FileGetAge(Name, CreateAge, LastAccessAge, LastWriteAge);
    FCreateAgesToCopy.Add(CreateAge);
    FLastAccessAgesToCopy.Add(LastAccessAge);
    FLastWriteAgesToCopy.Add(LastWriteAge);
    FFilesToCopyAge.Add(Dest);
  end;
  { 日時を複写するようになっていて }
  { 複写先のディレクトリに複写対象と同じディレクトリがあり }
  { 複写先のディレクトリの属性が複写対象と違えば }
  if FCopyAge
    and DirectoryExists(Dest)
    and (FileGetAttr(Dest) <> FileGetAttr(Name)) then
  begin
    FAttrsToCopy.Add(FileGetAttr(Name));
    FDirsToCopyAttr.Add(Dest);
  end;
end;

procedure TMainModule.OperateFilesToCopy;
var
  n: Integer;
  FilesToCopy, FilesToDest: String;
  SHFileOpStruct: TSHFileOpStruct;
begin
  { 複写対象がなければ }
  if FFilesToCopy.Count = 0 then
    Exit;
  { 複写対象を複写 }
  FilesToCopy := '';
  for n := 0 to FFilesToCopy.Count - 1 do
    FilesToCopy := FilesToCopy + FFilesToCopy.Strings[n] + #0;
  FilesToCopy := FilesToCopy + #0;
  FilesToDest := '';
  for n := 0 to FFilesToDest.Count - 1 do
    FilesToDest := FilesToDest + FFilesToDest.Strings[n] + #0;
  FilesToDest := FilesToDest + #0;
  with SHFileOpStruct do
  begin
    Wnd := GetDesktopWindow;
    if FMove then
      wFunc := FO_MOVE
    else
      wFunc := FO_COPY;
    pFrom := PChar(FilesToCopy);
    pTo := PChar(FilesToDest);
    fFlags := FOF_MULTIDESTFILES or FOF_ALLOWUNDO or FOF_NOCONFIRMMKDIR;
    if not FConfirm then
      fFlags := fFlags or FOF_NOCONFIRMATION;
    fAnyOperationsAborted := False;
    hNameMappings := nil;
    lpszProgressTitle := nil;
  end;
  if SHFileOperation(SHFileOpStruct) <> 0 then Abort;
end;

procedure TMainModule.CopyAgesToCopy;
var
  n: Integer;
  CreateAge, LastWriteAge, LastAccessAge: TDateTime;
  CopyCreateAge: Boolean;
begin
  { 日時を複写する指定がなければ }
  if not FCopyAge then
    Exit;
  { 同一ドライブ上で移動する場合は }
  if FMove
    and (ExtractFileDrive(FSourceDir) = ExtractFileDrive(FTargetDir)) then
    Exit;

  for n := 0 to FFilesToCopyAge.Count - 1 do
  begin
    { ディレクトリの日時を複写する指定がなければ }
    if DirectoryExists(FFilesToCopyAge.Strings[n]) and not FCopyDirAge then
      Continue;
    { 日時の複写先がネットワーク上のディレクトリなら }
    if DirectoryExists(FFilesToCopyAge.Strings[n]) and PathIsUNC(PChar(ExpandUNCFileName(FFilesToCopyAge.Strings[n]))) then
      { 強制的にフォルダの日時を変更する指定があり、NT 系 Windows なら }
      if FForce and OsIsWinNt then
        { 以下に続く }
      else
        Continue;
    { コピー先の作成日時がコピー元より古いとき日時をコピーしない指定なら }
    if not FCopyNewerCreateAge then
    begin
      FileGetAge(FFilesToCopyAge.Strings[n], CreateAge, LastAccessAge, LastWriteAge);
      CopyCreateAge := not (CreateAge < FCreateAgesToCopy.Items[n]);
    end
    else
      CopyCreateAge := True;
    { 複写先の日時を変更 }
    if not FileSetAge(FFilesToCopyAge.Strings[n],
      FCreateAgesToCopy.Items[n], FCopyCreateAge and CopyCreateAge,
      FLastAccessAgesToCopy.Items[n], FCopyLastAccessAge,
      FLastWriteAgesToCopy.Items[n], True) then
      if n < FFilesToCopyAge.Count - 1 then
        with CreateMessageDialog(QuotedStr(FFilesToCopyAge.Strings[n]) + ' の日時は変更できませんでした。'
          + #13#10#13#10 + '処理を中断しますか', mtError, [mbYes, mbNo]) do
        try
          ParentWindow := GetDesktopWindow;
          if ShowModal = mrYes then
            Exit;
        finally
          Free;
        end
      else
        with CreateMessageDialog(QuotedStr(FFilesToCopyAge.Strings[n]) + ' の日時は変更できませんでした。', mtError, [mbOK]) do
        try
          ParentWindow := GetDesktopWindow;
          ShowModal;
          Exit;
        finally
          Free;
        end;
  end;
end;

procedure TMainModule.CopyAttrsToCopy;
var
  n: Integer;
begin
  { 日時を複写する指定がなければ }
  if not FCopyAge then
    Exit;
  { 同一ドライブ上で移動する場合は }
  if FMove
    and (ExtractFileDrive(FSourceDir) = ExtractFileDrive(FTargetDir)) then
    Exit;

  for n := 0 to FDirsToCopyAttr.Count - 1 do
  begin
    { 複写先の属性を変更 }
    if FileSetAttr(FDirsToCopyAttr.Strings[n], FAttrsToCopy.Items[n]) <> 0 then
      if n < FDirsToCopyAttr.Count - 1 then
        with CreateMessageDialog(QuotedStr(FDirsToCopyAttr.Strings[n]) + ' の属性は変更できませんでした。'
          + #13#10#13#10 + '処理を中断しますか', mtError, [mbYes, mbNo]) do
        try
          ParentWindow := GetDesktopWindow;
          if ShowModal = mrYes then
            Exit;
        finally
          Free;
        end
      else
        with CreateMessageDialog(QuotedStr(FDirsToCopyAttr.Strings[n]) + ' の属性は変更できませんでした。', mtError, [mbOK]) do
        try
          ParentWindow := GetDesktopWindow;
          ShowModal;
          Exit;
        finally
          Free;
        end;
  end;
end;

procedure TMainModule.DoAdd;
begin
  try
    { 移動しない指定 }
    FMove := False;
    { 複写の対象になるファイル/ディレクトリを検索 }
    FindFilesToAdd;
    { 複写の対象のファイル/ディレクトリを複写 }
    OperateFilesToCopy;
    { 日時複写の対象のファイル/ディレクトリの日時を複写 }
    CopyAgesToCopy;
    { 属性複写の対象のファイル/ディレクトリの属性を複写 }
    CopyAttrsToCopy;
  except
    MessageDlg('ファイル/フォルダの追加を中断しました', mtError, [mbOK], 0);
  end;
end;

procedure TMainModule.FindFilesToAdd;
var
  n: Integer;
begin
  { 経過ダイアログを表示 }
  FProgressDlg.Show;
  FProgressDlg.Caption := '複写するﾌｧｲﾙ/ﾌｫﾙﾀﾞの検索';
  FProgressDlg.InitCount;
  { 変数を初期化 }
  FFilesToCopy.Clear;
  FFilesToDest.Clear;
  FCreateAgesToCopy.Clear;
  FLastWriteAgesToCopy.Clear;
  FFilesToCopyAge.Clear;
  { 複写の対象になるファイル/ディレクトリを検索 }
  for n := 0 to FSelectedFiles.Count - 1 do
  begin
    FSourceDir := ExtractFilePath(FSelectedFiles.Strings[n]);
    CheckFileToAdd(FSelectedFiles.Strings[n]);
    { 下位のディレクトリを検索する }
    if not DirectoryExists(FSelectedFiles.Strings[n]) then
      Continue;
    FSourceDir := FSelectedFiles.Strings[n];
    with FSearchFile do
    begin
      Directory := FSourceDir;
      SearchName := '*.*';
      Recursive := True;
      OnFind := FoundFileToAdd;
      if not Search then Abort;
    end;
  end;
  { 経過ダイアログを破棄 }
  FProgressDlg.Hide;
end;

procedure TMainModule.CheckFileToAdd(Name: String);
var
  Dest: String;
  CreateAge, LastAccessAge, LastWriteAge: TDateTime;
begin
  Dest := IncludeTrailingPathDelimiter(FTargetDir)
    + ExtractRelativePath(IncludeTrailingPathDelimiter(FSourceDir), Name);
  { 経過ダイアログへ表示 }
  FProgressDlg.SetFileName(Name);
  { 日時を複写するようになっていて }
  { 複写先のディレクトリに複写対象と同じディレクトリがあり }
  { 複写先のディレクトリの属性が複写対象と違えば }
  if FCopyAge
    and DirectoryExists(Dest)
    and (FileGetAttr(Dest) <> FileGetAttr(Name)) then
  begin
    FAttrsToCopy.Add(FileGetAttr(Name));
    FDirsToCopyAttr.Add(Dest);
  end;
  { 複写先のディレクトリに複写対象と同じディレクトリがあれば }
  if DirectoryExists(Dest) then
    Exit;
  { 複写先のディレクトリに複写対象と同じディレクトリがあり }
  { 複写先のファイルの属性が複写対象と同じなら }
  if DirectoryExists(Dest)
    and ((FileGetAttr(Dest) and faReadOnly) = (FileGetAttr(Name) and faReadOnly))
    and ((FileGetAttr(Dest) and faHidden) = (FileGetAttr(Name) and faHidden))
    and ((FileGetAttr(Dest) and faSysFile) = (FileGetAttr(Name) and faSysFile)) then
    Exit;
  { 複写先のディレクトリに複写対象と同じファイルがあり }
  { 複写先のファイルが複写対象より新しく }
  { 複写先のファイルの属性が複写対象と同じなら }
  if FileExists(Dest)
    and (FileAge(Dest) >= FileAge(Name))
    and ((FileGetAttr(Dest) and faReadOnly) = (FileGetAttr(Name) and faReadOnly))
    and ((FileGetAttr(Dest) and faHidden) = (FileGetAttr(Name) and faHidden))
    and ((FileGetAttr(Dest) and faSysFile) = (FileGetAttr(Name) and faSysFile)) then
    Exit;
  { 経過ダイアログへ表示 }
  FProgressDlg.AddCount;
  { リストに追加 }
  FFilesToCopy.Add(Name);
  FFilesToDest.Add(Dest);
  { 日時を複写するようになっていれば }
  if FCopyAge then
  begin
    FileGetAge(Name, CreateAge, LastAccessAge, LastWriteAge);
    FCreateAgesToCopy.Add(CreateAge);
    FLastAccessAgesToCopy.Add(LastAccessAge);
    FLastWriteAgesToCopy.Add(LastWriteAge);
    FFilesToCopyAge.Add(Dest);
  end;
end;

procedure TMainModule.FoundFileToAdd(Sender: TObject; Name: String;
  Data: TSearchFileData; var Continue: Boolean);
var
  Dest: String;
  n: Integer;
  CreateAge, LastWriteAge, LastAccessAge: TDateTime;
begin
  Application.ProcessMessages;
  if FProgressDlg.Stop then Continue := False;

  Dest := IncludeTrailingPathDelimiter(FTargetDir)
    + IncludeTrailingPathDelimiter(ExtractFileName(ExcludeTrailingPathDelimiter(FSourceDir)))
    + ExtractRelativePath(IncludeTrailingPathDelimiter(FSourceDir), Name);
  { 経過ダイアログへ表示 }
  FProgressDlg.SetFileName(Name);
  { 日時を複写するようになっていて }
  { 複写先のディレクトリに複写対象と同じディレクトリがあり }
  { 複写先のディレクトリの属性が複写対象と違えば }
  if FCopyAge
    and DirectoryExists(Dest)
    and ((FileGetAttr(Dest) and faReadOnly) = (FileGetAttr(Name) and faReadOnly))
    and ((FileGetAttr(Dest) and faHidden) = (FileGetAttr(Name) and faHidden))
    and ((FileGetAttr(Dest) and faSysFile) = (FileGetAttr(Name) and faSysFile)) then
  begin
    FAttrsToCopy.Add(FileGetAttr(Name));
    FDirsToCopyAttr.Add(Dest);
  end;
  { 複写先のディレクトリに複写対象と同じディレクトリがあれば }
  if DirectoryExists(Dest) then
    Exit;
  { 複写先のディレクトリに複写対象と同じファイルがあり }
  { 複写先のファイルが複写対象より新しく }
  { 複写先のファイルの属性が複写対象と同じなら }
  if FileExists(Dest)
    and (FileAge(Dest) >= FileAge(Name))
    and ((FileGetAttr(Dest) and faReadOnly) = (FileGetAttr(Name) and faReadOnly))
    and ((FileGetAttr(Dest) and faHidden) = (FileGetAttr(Name) and faHidden))
    and ((FileGetAttr(Dest) and faSysFile) = (FileGetAttr(Name) and faSysFile)) then
    Exit;
  { 経過ダイアログへ表示 }
  FProgressDlg.AddCount;
  { リストに追加 }
  for n := 0 to FFilesToCopy.Count - 1 do
    if Pos(FFilesToCopy.Strings[n], Name) = 1 then
      break;
  if n > FFilesToCopy.Count - 1 then  { 上位のディレクトリが複写対象になっていなければ }
  begin
    FFilesToCopy.Add(Name);
    FFilesToDest.Add(Dest);
  end;
  { 日時を複写するようになっていれば }
  if FCopyAge then
  begin
    FileGetAge(Name, CreateAge, LastAccessAge, LastWriteAge);
    FCreateAgesToCopy.Add(CreateAge);
    FLastAccessAgesToCopy.Add(LastAccessAge);
    FLastWriteAgesToCopy.Add(LastWriteAge);
    FFilesToCopyAge.Add(Dest);
  end;
end;

procedure TMainModule.DoUpdate;
begin
  try
    { コマンドラインで指定された複写/移動先の上のディレクトリを FTargetDir として }
    FTargetDir := ExtractFileDir(ExcludeTrailingBackslash(FTargetDir));
    { 移動しない指定 }
    FMove := False;
    { 削除の対象になるファイル/ディレクトリを検索 }
    FindFilesToDelete;
    { 複写の対象になるファイル/ディレクトリを検索 }
    FindFilesToAdd;
    { 削除の対象のファイル/ディレクトリを削除 }
    OperateFilesToDelete;
    { 複写の対象のファイル/ディレクトリを複写 }
    OperateFilesToCopy;
    { 日時複写の対象のファイル/ディレクトリの日時を複写 }
    CopyAgesToCopy;
    { 属性複写の対象のファイル/ディレクトリの属性を複写 }
    CopyAttrsToCopy;
  except
    MessageDlg('フォルダの更新を中断しました', mtError, [mbOK], 0);
  end;
end;

procedure TMainModule.FindFilesToDelete;
var
  n: Integer;
begin
  { 経過ダイアログを表示 }
  FProgressDlg.Show;
  FProgressDlg.Caption := '削除するﾌｧｲﾙ/ﾌｫﾙﾀﾞの検索';
  FProgressDlg.InitCount;
  { 削除の対象になるファイル/ディレクトリを検索 }
  FFilesToDelete.Clear;
  for n := 0 to FSelectedFiles.Count - 1 do
  begin
    if not DirectoryExists(FSelectedFiles.Strings[n]) then
      Continue;
    FSourceDir := FSelectedFiles.Strings[n];
    with FSearchFile do
    begin
      Directory := IncludeTrailingPathDelimiter(FTargetDir)
        + ExtractFileName(ExcludeTrailingPathDelimiter(FSourceDir));
      SearchName := '*.*';
      Recursive := True;
      OnFind := FoundFileToDelete;
      if not Search then Abort;
    end;
  end;
  { 経過ダイアログを破棄 }
  FProgressDlg.Hide;
end;

procedure TMainModule.FoundFileToDelete(Sender: TObject; Name: String;
  Data: TSearchFileData; var Continue: Boolean);
var
  Source: String;
  n: Integer;
begin
  Application.ProcessMessages;
  if FProgressDlg.Stop then Continue := False;

  Source := IncludeTrailingPathDelimiter(FSourceDir)
    + ExtractRelativePath(IncludeTrailingPathDelimiter(IncludeTrailingPathDelimiter(FTargetDir)
    + ExtractFileName(ExcludeTrailingPathDelimiter(FSourceDir))), Name);
  { 経過ダイアログへ表示 }
  FProgressDlg.SetFileName(Name);
  { 複写元のディレクトリに複写対象と同じファイルがあれば }
  if FileExists(Source) then
    Exit;
  { 複写元のディレクトリに複写対象と同じディレクトリがあれば }
  if DirectoryExists(Source) then
    Exit;
  { 上位のディレクトリが削除対象になっていれば }
  for n := 0 to FFilesToDelete.Count - 1 do
    if Pos(FFilesToDelete.Strings[n], Name) = 1 then
      Exit;
  { 経過ダイアログへ表示 }
  FProgressDlg.AddCount;
  { 削除対象に追加 }
  FFilesToDelete.Add(Name);
end;

procedure TMainModule.OperateFilesToDelete;
var
  FilesToDelete: String;
  n: Integer;
  SHFileOpStruct: TSHFileOpStruct;
begin
  { 削除対象がなければ }
  if FFilesToDelete.Count = 0 then
    Exit;
  { 削除対象を削除 }
  FilesToDelete := '';
  for n := 0 to FFilesToDelete.Count - 1 do
    FilesToDelete := FilesToDelete + FFilesToDelete.Strings[n] + #0;
  FilesToDelete := FilesToDelete + #0;
  with SHFileOpStruct do
  begin
    Wnd := GetDesktopWindow;
    wFunc := FO_DELETE;
    pFrom := PChar(FilesToDelete);
    pTo := nil;
    fFlags := FOF_ALLOWUNDO;
    if not FConfirm then
      fFlags := fFlags or FOF_NOCONFIRMATION;
    fAnyOperationsAborted := False;
    hNameMappings := nil;
    lpszProgressTitle := nil;
  end;
  if SHFileOperation(SHFileOpStruct) <> 0 then Abort;
end;

procedure TMainModule.DoBackup;
begin
  try
    { 移動しない指定 }
    FMove := False;
    { 削除の対象になるファイル/ディレクトリを検索 }
    FindFilesToDelete;
    { 複写の対象になるファイル/ディレクトリを検索 }
    FindFilesToAdd;
    { 削除の対象のファイル/ディレクトリを削除 }
    OperateFilesToDelete;
    { 複写の対象のファイル/ディレクトリを複写 }
    OperateFilesToCopy;
    { 日時複写の対象のファイル/ディレクトリの日時を複写 }
    CopyAgesToCopy;
  except
    MessageDlg('ファイル/フォルダのバックアップを中断しました', mtError, [mbOK], 0);
  end;
end;

procedure TMainModule.DoCopyAge;
begin
  try
    { 移動しない指定 }
    FMove := False;
    { 複写の対象になるファイル/ディレクトリを検索 }
    FindFilesToCopy;
    { 日時複写の対象のファイル/ディレクトリの日時を複写 }
    CopyAgesToCopy;
    { 属性複写の対象のファイル/ディレクトリの属性を複写 }
    CopyAttrsToCopy;
  except
    MessageDlg('ファイル/フォルダのコピーを中断しました', mtError, [mbOK], 0);
  end;
end;

end.
