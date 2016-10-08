program FileCopy;

uses
  Forms,
  Main in 'Main.pas' {MainModule},
  Progress in 'Progress.pas' {ProgressDlg};

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TMainModule, MainModule);
  MainModule.RunCommand;
end.
