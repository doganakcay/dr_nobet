program Dr_nobet;

uses
  Forms,
  udr_cnobet in 'udr_cnobet.pas' {FRM_DR_CNOBET},
  udr_nob_sayi in 'udr_nob_sayi.pas' {frmIstatistik};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFRM_DR_CNOBET, FRM_DR_CNOBET);
  Application.CreateForm(TfrmIstatistik, frmIstatistik);
  Application.Run;
end.
