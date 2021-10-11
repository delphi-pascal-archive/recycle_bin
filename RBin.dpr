program RBin;

uses
  Forms,
  Unit1 in 'Unit1.pas' {frmRecycleBin};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmRecycleBin, frmRecycleBin);
  Application.Run;
end.
