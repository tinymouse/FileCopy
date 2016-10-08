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
    { Private �錾 }
    FSelectedFiles: TStringList;  { �I�����ꂽ�t�@�C��/�f�B���N�g���̃p�X }
    FSelectedDirs: Boolean;  { �I�����ꂽ�f�B���N�g�������邩 }
    FSourceDir: String;
    FTargetDir: String;  { ����/�ړ�/�ǉ�/�X�V��̃f�B���N�g���̃p�X }
    FFilesToCopy: TStringList;  { ����/�ړ��̑ΏۂƂȂ�t�@�C��/�f�B���N�g���̃p�X }
    FFilesToDest: TStringList;
    FFilesToDelete: TStringList;  { �폜�̑ΏۂƂȂ�t�@�C��/�f�B���N�g���̃p�X }
    FCreateAgesToCopy: TDateTimeList;  { �������ʂ̑ΏۂƂȂ�t�@�C��/�f�B���N�g���̍쐬���� }
    FLastAccessAgesToCopy: TDateTimeList;    { �������ʂ̑ΏۂƂȂ�t�@�C��/�f�B���N�g���̃A�N�Z�X���� }
    FLastWriteAgesToCopy: TDateTimeList;  { �������ʂ̑ΏۂƂȂ�t�@�C��/�f�B���N�g���̍X�V���� }
    FFilesToCopyAge: TStringList;  { �������ʂ̑ΏۂƂȂ�t�@�C��/�f�B���N�g���̕���/�ړ���̃p�X }
    FAttrsToCopy: TIntegerList;  { �������ʂ̑ΏۂƂȂ�f�B���N�g���̑��� }
    FDirsToCopyAttr: TStringList;  { �������ʂ̑ΏۂƂȂ�f�B���N�g���̕���/�ړ���̃p�X }
    FMove: Boolean;  { �ړ����ۂ� }
    FCopyAge: Boolean;  { ����/�����𕡎ʂ��邩 }
    FConfirm: Boolean;  { �㏑�m�F�����邩 }
    FForce: Boolean;  { �����I�Ƀt�H���_�̓�����ύX���邩 }
    FCopyCreateAge: Boolean;  { �쐬�����͕ێ����邩 }
    FCopyNewerCreateAge: Boolean;  { �R�s�[��̍쐬�������R�s�[�����Â��Ƃ��������R�s�[���邩 }
    FCopyLastAccessAge: Boolean;  { �A�N�Z�X�������ێ����邩 }
    FCopyDirAge: Boolean;  { �f�B���N�g���̓������ێ����邩 }
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
    { Public �錾 }
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
  { �ϐ��������� }
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
  { �ϐ������ }
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
    MessageDlg('�R�}���h�w�肪�Ԉ���Ă��܂�', mtError, [mbOK], 0);
    Exit;
  end;

  { ����/�ړ����ƕ���/�ړ�����R�}���h���C���ϐ�����擾 }
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
  { �擾���ꂽ�t�@�C��/�f�B���N�g���̐��� 0 �Ȃ� }
  if FSelectedFiles.Count = 0 then
  begin
    MessageDlg('����/�ړ�����t�@�C��/�f�B���N�g��������܂���', mtError, [mbOK], 0);
    Exit;
  end;

  { �I�v�V�����w����R�}���h���C���ϐ�����擾 }
  FCopyAge := FindCmdLineSwitch('copyage');
  FConfirm := not FindCmdLineSwitch('noconfirm');
  FForce := FindCmdLineSwitch('force');
  FCopyCreateAge := not FindCmdLineSwitch('notcopycreateage');
  FCopyNewerCreateAge := not FindCmdLineSwitch('notcopynewercreateage');
  FCopyLastAccessAge := FindCmdLineSwitch('copylastaccessage');
  FCopyDirAge := not FindCmdLineSwitch('notcopydirage');

  { �f�B���N�g�����I������Ă��āA}
  { �����𕡎ʂ���悤�ɂȂ��Ă��āA}
  { ���ʐ悪�l�b�g���[�N��̃f�B���N�g���Ȃ� }
  if FSelectedDirs
    and FCopyAge
    and PathIsUNC(PChar(ExpandUNCFileName(FTargetDir))) then
  begin
    { �����I�Ƀt�H���_�̓�����ύX����w�肪����ANT �n Windows �Ȃ� }
    if FForce and OsIsWinNt then
      { �������Ȃ� }
    else
      with CreateMessageDialog(QuotedStr(FTargetDir) + ' �̓l�b�g���[�N��̃t�H���_�Ȃ̂ŁA'
        + '�t�H���_�̓����͕ێ����邱�Ƃ��ł��܂���'
        + #13#10#13#10 + '�����𑱂��܂���', mtError, [mbYes, mbNo]) do
      try
        ParentWindow := GetDesktopWindow;
        if ShowModal = mrNo then
          Exit;
      finally
        Free;
      end;
  end;

  { �w�߂��R�}���h���C���ϐ�����擾 }
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
    MessageDlg('�R�}���h�w�肪�Ԉ���Ă��܂�', mtError, [mbOK], 0);
end;

procedure TMainModule.DoCopy;
begin
  try
    { �ړ����Ȃ��w�� }
    FMove := False;
    { ���ʂ̑ΏۂɂȂ�t�@�C��/�f�B���N�g�������� }
    FindFilesToCopy;
    { ���ʂ̑Ώۂ̃t�@�C��/�f�B���N�g���𕡎� }
    OperateFilesToCopy;
    { �������ʂ̑Ώۂ̃t�@�C��/�f�B���N�g���̓����𕡎� }
    CopyAgesToCopy;
    { �������ʂ̑Ώۂ̃t�@�C��/�f�B���N�g���̑����𕡎� }
    CopyAttrsToCopy;
  except
    MessageDlg('�t�@�C��/�t�H���_�̃R�s�[�𒆒f���܂���', mtError, [mbOK], 0);
  end;
end;

procedure TMainModule.DoMove;
begin
  try
    { �ړ�����w�� }
    FMove := True;
    { �ړ��̑ΏۂɂȂ�t�@�C��/�f�B���N�g�������� }
    FindFilesToCopy;
    { �ړ��̑Ώۂ̃t�@�C��/�f�B���N�g�����ړ� }
    OperateFilesToCopy;
    { �������ʂ̑Ώۂ̃t�@�C��/�f�B���N�g���̓����𕡎� }
    CopyAgesToCopy;
    { �������ʂ̑Ώۂ̃t�@�C��/�f�B���N�g���̑����𕡎� }
    CopyAttrsToCopy;
  except
    MessageDlg('�t�@�C��/�t�H���_�̈ړ��𒆒f���܂���', mtError, [mbOK], 0);
  end;
end;

procedure TMainModule.FindFilesToCopy;
var
  n: Integer;
begin
  { �o�߃_�C�A���O��\�� }
  FProgressDlg.Show;
  if FMove then
    FProgressDlg.Caption := '�ړ�����̧��/̫��ނ̌���'
  else
    FProgressDlg.Caption := '���ʂ���̧��/̫��ނ̌���';
  FProgressDlg.InitCount;
  { �ϐ��������� }
  FFilesToCopy.Clear;
  FFilesToDest.Clear;
  FCreateAgesToCopy.Clear;
  FLastWriteAgesToCopy.Clear;
  FFilesToCopyAge.Clear;
  { ���ʂ̑ΏۂɂȂ�t�@�C��/�f�B���N�g�������� }
  for n := 0 to FSelectedFiles.Count - 1 do
  begin
    FSourceDir := ExtractFilePath(FSelectedFiles.Strings[n]);
    CheckFileToCopy(FSelectedFiles.Strings[n]);
    { ���ʂ̃f�B���N�g������������ }
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
  { �o�߃_�C�A���O��j�� }
  FProgressDlg.Hide;
end;

procedure TMainModule.CheckFileToCopy(Name: String);
var
  Dest: String;
  CreateAge, LastAccessAge, LastWriteAge: TDateTime;
begin
  Dest := IncludeTrailingPathDelimiter(FTargetDir)
    + ExtractRelativePath(IncludeTrailingPathDelimiter(FSourceDir), Name);
  { �o�߃_�C�A���O�֕\�� }
  FProgressDlg.SetFileName(Name);
  FProgressDlg.AddCount;
  { ���ʑΏۂɒǉ� }
  FFilesToCopy.Add(Name);
  FFilesToDest.Add(Dest);
  { �����𕡎ʂ���悤�ɂȂ��Ă���� }
  if FCopyAge then
  begin
    FileGetAge(Name, CreateAge, LastAccessAge, LastWriteAge);
    FCreateAgesToCopy.Add(CreateAge);
    FLastAccessAgesToCopy.Add(LastAccessAge);
    FLastWriteAgesToCopy.Add(LastWriteAge);
    FFilesToCopyAge.Add(Dest);
  end;
  { �����𕡎ʂ���悤�ɂȂ��Ă��� }
  { ���ʐ�̃f�B���N�g���ɕ��ʑΏۂƓ����f�B���N�g�������� }
  { ���ʐ�̃f�B���N�g���̑��������ʑΏۂƈႦ�� }
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
  { �o�߃_�C�A���O�֕\�� }
  FProgressDlg.SetFileName(Name);
  FProgressDlg.AddCount;
  { ���ʑΏۂɒǉ� }
  for n := 0 to FFilesToCopy.Count - 1 do
    if Pos(FFilesToCopy.Strings[n], Name) = 1 then
      break;
  if n > FFilesToCopy.Count - 1 then  { ��ʂ̃f�B���N�g�������ʑΏۂɂȂ��Ă��Ȃ���� }
  begin
    FFilesToCopy.Add(Name);
    FFilesToDest.Add(Dest);
  end;
  { �����𕡎ʂ���悤�ɂȂ��Ă���� }
  if FCopyAge then
  begin
    FileGetAge(Name, CreateAge, LastAccessAge, LastWriteAge);
    FCreateAgesToCopy.Add(CreateAge);
    FLastAccessAgesToCopy.Add(LastAccessAge);
    FLastWriteAgesToCopy.Add(LastWriteAge);
    FFilesToCopyAge.Add(Dest);
  end;
  { �����𕡎ʂ���悤�ɂȂ��Ă��� }
  { ���ʐ�̃f�B���N�g���ɕ��ʑΏۂƓ����f�B���N�g�������� }
  { ���ʐ�̃f�B���N�g���̑��������ʑΏۂƈႦ�� }
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
  { ���ʑΏۂ��Ȃ���� }
  if FFilesToCopy.Count = 0 then
    Exit;
  { ���ʑΏۂ𕡎� }
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
  { �����𕡎ʂ���w�肪�Ȃ���� }
  if not FCopyAge then
    Exit;
  { ����h���C�u��ňړ�����ꍇ�� }
  if FMove
    and (ExtractFileDrive(FSourceDir) = ExtractFileDrive(FTargetDir)) then
    Exit;

  for n := 0 to FFilesToCopyAge.Count - 1 do
  begin
    { �f�B���N�g���̓����𕡎ʂ���w�肪�Ȃ���� }
    if DirectoryExists(FFilesToCopyAge.Strings[n]) and not FCopyDirAge then
      Continue;
    { �����̕��ʐ悪�l�b�g���[�N��̃f�B���N�g���Ȃ� }
    if DirectoryExists(FFilesToCopyAge.Strings[n]) and PathIsUNC(PChar(ExpandUNCFileName(FFilesToCopyAge.Strings[n]))) then
      { �����I�Ƀt�H���_�̓�����ύX����w�肪����ANT �n Windows �Ȃ� }
      if FForce and OsIsWinNt then
        { �ȉ��ɑ��� }
      else
        Continue;
    { �R�s�[��̍쐬�������R�s�[�����Â��Ƃ��������R�s�[���Ȃ��w��Ȃ� }
    if not FCopyNewerCreateAge then
    begin
      FileGetAge(FFilesToCopyAge.Strings[n], CreateAge, LastAccessAge, LastWriteAge);
      CopyCreateAge := not (CreateAge < FCreateAgesToCopy.Items[n]);
    end
    else
      CopyCreateAge := True;
    { ���ʐ�̓�����ύX }
    if not FileSetAge(FFilesToCopyAge.Strings[n],
      FCreateAgesToCopy.Items[n], FCopyCreateAge and CopyCreateAge,
      FLastAccessAgesToCopy.Items[n], FCopyLastAccessAge,
      FLastWriteAgesToCopy.Items[n], True) then
      if n < FFilesToCopyAge.Count - 1 then
        with CreateMessageDialog(QuotedStr(FFilesToCopyAge.Strings[n]) + ' �̓����͕ύX�ł��܂���ł����B'
          + #13#10#13#10 + '�����𒆒f���܂���', mtError, [mbYes, mbNo]) do
        try
          ParentWindow := GetDesktopWindow;
          if ShowModal = mrYes then
            Exit;
        finally
          Free;
        end
      else
        with CreateMessageDialog(QuotedStr(FFilesToCopyAge.Strings[n]) + ' �̓����͕ύX�ł��܂���ł����B', mtError, [mbOK]) do
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
  { �����𕡎ʂ���w�肪�Ȃ���� }
  if not FCopyAge then
    Exit;
  { ����h���C�u��ňړ�����ꍇ�� }
  if FMove
    and (ExtractFileDrive(FSourceDir) = ExtractFileDrive(FTargetDir)) then
    Exit;

  for n := 0 to FDirsToCopyAttr.Count - 1 do
  begin
    { ���ʐ�̑�����ύX }
    if FileSetAttr(FDirsToCopyAttr.Strings[n], FAttrsToCopy.Items[n]) <> 0 then
      if n < FDirsToCopyAttr.Count - 1 then
        with CreateMessageDialog(QuotedStr(FDirsToCopyAttr.Strings[n]) + ' �̑����͕ύX�ł��܂���ł����B'
          + #13#10#13#10 + '�����𒆒f���܂���', mtError, [mbYes, mbNo]) do
        try
          ParentWindow := GetDesktopWindow;
          if ShowModal = mrYes then
            Exit;
        finally
          Free;
        end
      else
        with CreateMessageDialog(QuotedStr(FDirsToCopyAttr.Strings[n]) + ' �̑����͕ύX�ł��܂���ł����B', mtError, [mbOK]) do
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
    { �ړ����Ȃ��w�� }
    FMove := False;
    { ���ʂ̑ΏۂɂȂ�t�@�C��/�f�B���N�g�������� }
    FindFilesToAdd;
    { ���ʂ̑Ώۂ̃t�@�C��/�f�B���N�g���𕡎� }
    OperateFilesToCopy;
    { �������ʂ̑Ώۂ̃t�@�C��/�f�B���N�g���̓����𕡎� }
    CopyAgesToCopy;
    { �������ʂ̑Ώۂ̃t�@�C��/�f�B���N�g���̑����𕡎� }
    CopyAttrsToCopy;
  except
    MessageDlg('�t�@�C��/�t�H���_�̒ǉ��𒆒f���܂���', mtError, [mbOK], 0);
  end;
end;

procedure TMainModule.FindFilesToAdd;
var
  n: Integer;
begin
  { �o�߃_�C�A���O��\�� }
  FProgressDlg.Show;
  FProgressDlg.Caption := '���ʂ���̧��/̫��ނ̌���';
  FProgressDlg.InitCount;
  { �ϐ��������� }
  FFilesToCopy.Clear;
  FFilesToDest.Clear;
  FCreateAgesToCopy.Clear;
  FLastWriteAgesToCopy.Clear;
  FFilesToCopyAge.Clear;
  { ���ʂ̑ΏۂɂȂ�t�@�C��/�f�B���N�g�������� }
  for n := 0 to FSelectedFiles.Count - 1 do
  begin
    FSourceDir := ExtractFilePath(FSelectedFiles.Strings[n]);
    CheckFileToAdd(FSelectedFiles.Strings[n]);
    { ���ʂ̃f�B���N�g������������ }
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
  { �o�߃_�C�A���O��j�� }
  FProgressDlg.Hide;
end;

procedure TMainModule.CheckFileToAdd(Name: String);
var
  Dest: String;
  CreateAge, LastAccessAge, LastWriteAge: TDateTime;
begin
  Dest := IncludeTrailingPathDelimiter(FTargetDir)
    + ExtractRelativePath(IncludeTrailingPathDelimiter(FSourceDir), Name);
  { �o�߃_�C�A���O�֕\�� }
  FProgressDlg.SetFileName(Name);
  { �����𕡎ʂ���悤�ɂȂ��Ă��� }
  { ���ʐ�̃f�B���N�g���ɕ��ʑΏۂƓ����f�B���N�g�������� }
  { ���ʐ�̃f�B���N�g���̑��������ʑΏۂƈႦ�� }
  if FCopyAge
    and DirectoryExists(Dest)
    and (FileGetAttr(Dest) <> FileGetAttr(Name)) then
  begin
    FAttrsToCopy.Add(FileGetAttr(Name));
    FDirsToCopyAttr.Add(Dest);
  end;
  { ���ʐ�̃f�B���N�g���ɕ��ʑΏۂƓ����f�B���N�g��������� }
  if DirectoryExists(Dest) then
    Exit;
  { ���ʐ�̃f�B���N�g���ɕ��ʑΏۂƓ����f�B���N�g�������� }
  { ���ʐ�̃t�@�C���̑��������ʑΏۂƓ����Ȃ� }
  if DirectoryExists(Dest)
    and ((FileGetAttr(Dest) and faReadOnly) = (FileGetAttr(Name) and faReadOnly))
    and ((FileGetAttr(Dest) and faHidden) = (FileGetAttr(Name) and faHidden))
    and ((FileGetAttr(Dest) and faSysFile) = (FileGetAttr(Name) and faSysFile)) then
    Exit;
  { ���ʐ�̃f�B���N�g���ɕ��ʑΏۂƓ����t�@�C�������� }
  { ���ʐ�̃t�@�C�������ʑΏۂ��V���� }
  { ���ʐ�̃t�@�C���̑��������ʑΏۂƓ����Ȃ� }
  if FileExists(Dest)
    and (FileAge(Dest) >= FileAge(Name))
    and ((FileGetAttr(Dest) and faReadOnly) = (FileGetAttr(Name) and faReadOnly))
    and ((FileGetAttr(Dest) and faHidden) = (FileGetAttr(Name) and faHidden))
    and ((FileGetAttr(Dest) and faSysFile) = (FileGetAttr(Name) and faSysFile)) then
    Exit;
  { �o�߃_�C�A���O�֕\�� }
  FProgressDlg.AddCount;
  { ���X�g�ɒǉ� }
  FFilesToCopy.Add(Name);
  FFilesToDest.Add(Dest);
  { �����𕡎ʂ���悤�ɂȂ��Ă���� }
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
  { �o�߃_�C�A���O�֕\�� }
  FProgressDlg.SetFileName(Name);
  { �����𕡎ʂ���悤�ɂȂ��Ă��� }
  { ���ʐ�̃f�B���N�g���ɕ��ʑΏۂƓ����f�B���N�g�������� }
  { ���ʐ�̃f�B���N�g���̑��������ʑΏۂƈႦ�� }
  if FCopyAge
    and DirectoryExists(Dest)
    and ((FileGetAttr(Dest) and faReadOnly) = (FileGetAttr(Name) and faReadOnly))
    and ((FileGetAttr(Dest) and faHidden) = (FileGetAttr(Name) and faHidden))
    and ((FileGetAttr(Dest) and faSysFile) = (FileGetAttr(Name) and faSysFile)) then
  begin
    FAttrsToCopy.Add(FileGetAttr(Name));
    FDirsToCopyAttr.Add(Dest);
  end;
  { ���ʐ�̃f�B���N�g���ɕ��ʑΏۂƓ����f�B���N�g��������� }
  if DirectoryExists(Dest) then
    Exit;
  { ���ʐ�̃f�B���N�g���ɕ��ʑΏۂƓ����t�@�C�������� }
  { ���ʐ�̃t�@�C�������ʑΏۂ��V���� }
  { ���ʐ�̃t�@�C���̑��������ʑΏۂƓ����Ȃ� }
  if FileExists(Dest)
    and (FileAge(Dest) >= FileAge(Name))
    and ((FileGetAttr(Dest) and faReadOnly) = (FileGetAttr(Name) and faReadOnly))
    and ((FileGetAttr(Dest) and faHidden) = (FileGetAttr(Name) and faHidden))
    and ((FileGetAttr(Dest) and faSysFile) = (FileGetAttr(Name) and faSysFile)) then
    Exit;
  { �o�߃_�C�A���O�֕\�� }
  FProgressDlg.AddCount;
  { ���X�g�ɒǉ� }
  for n := 0 to FFilesToCopy.Count - 1 do
    if Pos(FFilesToCopy.Strings[n], Name) = 1 then
      break;
  if n > FFilesToCopy.Count - 1 then  { ��ʂ̃f�B���N�g�������ʑΏۂɂȂ��Ă��Ȃ���� }
  begin
    FFilesToCopy.Add(Name);
    FFilesToDest.Add(Dest);
  end;
  { �����𕡎ʂ���悤�ɂȂ��Ă���� }
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
    { �R�}���h���C���Ŏw�肳�ꂽ����/�ړ���̏�̃f�B���N�g���� FTargetDir �Ƃ��� }
    FTargetDir := ExtractFileDir(ExcludeTrailingBackslash(FTargetDir));
    { �ړ����Ȃ��w�� }
    FMove := False;
    { �폜�̑ΏۂɂȂ�t�@�C��/�f�B���N�g�������� }
    FindFilesToDelete;
    { ���ʂ̑ΏۂɂȂ�t�@�C��/�f�B���N�g�������� }
    FindFilesToAdd;
    { �폜�̑Ώۂ̃t�@�C��/�f�B���N�g�����폜 }
    OperateFilesToDelete;
    { ���ʂ̑Ώۂ̃t�@�C��/�f�B���N�g���𕡎� }
    OperateFilesToCopy;
    { �������ʂ̑Ώۂ̃t�@�C��/�f�B���N�g���̓����𕡎� }
    CopyAgesToCopy;
    { �������ʂ̑Ώۂ̃t�@�C��/�f�B���N�g���̑����𕡎� }
    CopyAttrsToCopy;
  except
    MessageDlg('�t�H���_�̍X�V�𒆒f���܂���', mtError, [mbOK], 0);
  end;
end;

procedure TMainModule.FindFilesToDelete;
var
  n: Integer;
begin
  { �o�߃_�C�A���O��\�� }
  FProgressDlg.Show;
  FProgressDlg.Caption := '�폜����̧��/̫��ނ̌���';
  FProgressDlg.InitCount;
  { �폜�̑ΏۂɂȂ�t�@�C��/�f�B���N�g�������� }
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
  { �o�߃_�C�A���O��j�� }
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
  { �o�߃_�C�A���O�֕\�� }
  FProgressDlg.SetFileName(Name);
  { ���ʌ��̃f�B���N�g���ɕ��ʑΏۂƓ����t�@�C��������� }
  if FileExists(Source) then
    Exit;
  { ���ʌ��̃f�B���N�g���ɕ��ʑΏۂƓ����f�B���N�g��������� }
  if DirectoryExists(Source) then
    Exit;
  { ��ʂ̃f�B���N�g�����폜�ΏۂɂȂ��Ă���� }
  for n := 0 to FFilesToDelete.Count - 1 do
    if Pos(FFilesToDelete.Strings[n], Name) = 1 then
      Exit;
  { �o�߃_�C�A���O�֕\�� }
  FProgressDlg.AddCount;
  { �폜�Ώۂɒǉ� }
  FFilesToDelete.Add(Name);
end;

procedure TMainModule.OperateFilesToDelete;
var
  FilesToDelete: String;
  n: Integer;
  SHFileOpStruct: TSHFileOpStruct;
begin
  { �폜�Ώۂ��Ȃ���� }
  if FFilesToDelete.Count = 0 then
    Exit;
  { �폜�Ώۂ��폜 }
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
    { �ړ����Ȃ��w�� }
    FMove := False;
    { �폜�̑ΏۂɂȂ�t�@�C��/�f�B���N�g�������� }
    FindFilesToDelete;
    { ���ʂ̑ΏۂɂȂ�t�@�C��/�f�B���N�g�������� }
    FindFilesToAdd;
    { �폜�̑Ώۂ̃t�@�C��/�f�B���N�g�����폜 }
    OperateFilesToDelete;
    { ���ʂ̑Ώۂ̃t�@�C��/�f�B���N�g���𕡎� }
    OperateFilesToCopy;
    { �������ʂ̑Ώۂ̃t�@�C��/�f�B���N�g���̓����𕡎� }
    CopyAgesToCopy;
  except
    MessageDlg('�t�@�C��/�t�H���_�̃o�b�N�A�b�v�𒆒f���܂���', mtError, [mbOK], 0);
  end;
end;

procedure TMainModule.DoCopyAge;
begin
  try
    { �ړ����Ȃ��w�� }
    FMove := False;
    { ���ʂ̑ΏۂɂȂ�t�@�C��/�f�B���N�g�������� }
    FindFilesToCopy;
    { �������ʂ̑Ώۂ̃t�@�C��/�f�B���N�g���̓����𕡎� }
    CopyAgesToCopy;
    { �������ʂ̑Ώۂ̃t�@�C��/�f�B���N�g���̑����𕡎� }
    CopyAttrsToCopy;
  except
    MessageDlg('�t�@�C��/�t�H���_�̃R�s�[�𒆒f���܂���', mtError, [mbOK], 0);
  end;
end;

end.
