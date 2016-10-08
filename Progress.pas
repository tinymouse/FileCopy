unit Progress;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ComCtrls;

type
  TProgressDlg = class(TForm)
    lblFileName: TLabel;
    txtFileName: TLabel;
    lblCount: TLabel;
    txtCount: TLabel;
    btnCancel: TButton;
    aniFindFile: TAnimate;
    procedure FormShow(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
  private
    { Private 宣言 }
    FAll, FCount: Integer;
    FStop: Boolean;
  public
    { Public 宣言 }
    procedure InitCount;
    procedure SetFileName(Name: String);
    procedure AddCount;
    property Stop: Boolean read FStop;
  end;

var
  ProgressDlg: TProgressDlg;

implementation

{$R *.dfm}

uses
  FileCtrl, ShellApi, StrUtils, Utils;

procedure TProgressDlg.InitCount;
begin
  FAll := 0;
  FCount := 0;
  FStop := False;
end;

procedure TProgressDlg.SetFileName(Name: String);
begin
  FAll := FAll + 1;
  txtFileName.Caption := MinimizeName(Name, txtFileName.Canvas, txtFileName.Width);
  txtCount.Caption := Format('%d/%d', [FCount, FAll]);
  Refresh;
end;

procedure TProgressDlg.AddCount;
begin
  FCount := FCount + 1;
  txtCount.Caption := Format('%d/%d', [FCount, FAll]);
  txtCount.Refresh;
end;

procedure TProgressDlg.FormShow(Sender: TObject);
begin
  ShowWindow(Application.Handle, SW_HIDE);  { タスクバーにアイコンを表示させない }
end;

procedure TProgressDlg.btnCancelClick(Sender: TObject);
begin
  FStop := True;
end;

end.
