unit udr_nob_sayi;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, DBAccess, Ora, MemDS, Grids, DBGrids, ExtCtrls;

type
  TfrmIstatistik = class(TForm)
    Panel1: TPanel;
    DBGrid1: TDBGrid;
    QISTATISTIK: TOraQuery;
    DSISTATISTIK: TOraDataSource;
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmIstatistik: TfrmIstatistik;

implementation

uses udr_cnobet;

{$R *.dfm}

procedure TfrmIstatistik.FormActivate(Sender: TObject);
begin
TOP:=0;
LEFT:=(Screen.Width DIV 2)-(Width div 2);
end;

end.
