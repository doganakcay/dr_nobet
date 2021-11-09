unit udr_Cnobet;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, dxSkinsCore, dxSkinsDefaultPainters, dxSkinscxPCPainter, cxStyles,
  cxCustomData, cxGraphics, cxFilter, cxData, cxDataStorage, cxEdit, DB,
  cxDBData, cxGridCustomTableView, cxGridTableView, cxGridDBTableView,
  cxGridLevel, cxClasses, cxControls, cxGridCustomView, cxGrid, DBAccess, Ora,
  ExtCtrls, MemDS, cxDBLookupComboBox, DBCtrls, cxContainer, cxTextEdit,
  cxMaskEdit, cxDropDownEdit, cxLookupEdit, cxDBLookupEdit, cxButtonEdit,
  StdCtrls, Grids, DBGrids, RzPanel, RzSplit, RzGroupBar, RzTabs, RzLabel,
  RzButton, frxClass, frxDBSet,frxexportxls,frxExportPDF,frxExportMail
  ,frxExportRTF,frxExportImage, frxDMPExport, frxExportHTML,
  frxExportXML, frxExportText, frxExportCSV, frxExportODF, Menus,Printers,
  frxExportTXT;



type
  TFRM_DR_CNOBET = class(TForm)

    QDRLOOP: TOraQuery;
    dsloop: TOraDataSource;
    QTABLO: TOraQuery;
    DSTABLO: TOraDataSource;
    Panel2: TPanel;
    Panel3: TPanel;
    
    ctarih: TcxComboBox;
    RzLabel1: TRzLabel;
    RzBitBtn1: TRzBitBtn;
    takvim_olustur: TOraStoredProc;
    RzBitBtn2: TRzBitBtn;
    DSLOOPCMPTZ: TOraDataSource;
    QDRLOOPCMPTZ: TOraQuery;
    sql_getir: TOraStoredProc;
    Edit1: TEdit;
    cxLookupComboBox1: TcxLookupComboBox;
    SQL_INSER: TOraSQL;
    PopupMenu1: TPopupMenu;
    Ekle1: TMenuItem;
    Sil1: TMenuItem;
    frxReport1: TfrxReport;
    frxDBDataset1: TfrxDBDataset;
    frxPDFExport1: TfrxPDFExport;
    frxJPEGExport1: TfrxJPEGExport;
    frxBMPExport1: TfrxBMPExport;
    frxTXTExport1: TfrxTXTExport;
    frxXLSExport1: TfrxXLSExport;
    frxRTFExport1: TfrxRTFExport;
    AKGUN: TOraSession;
    QBRANS: TOraQuery;
    DSBRANS: TOraDataSource;
    EBRANSI: TcxLookupComboBox;
    RzLabel2: TRzLabel;
    RzBitBtn3: TRzBitBtn;
    DBGrid1: TDBGrid;
    QETIKET: TOraQuery;
    RzLabel3: TRzLabel;
    SQL_GETIR_TEK: TOraStoredProc;
    QTABLO_TEK: TOraQuery;
    DS_TABLO_TEK: TOraDataSource;
    DBGrid2: TDBGrid;
    procedure FormCreate(Sender: TObject);
    Procedure tabloloriac;
    Procedure tabloloriac_TEK;
    procedure ctarihPropertiesChange(Sender: TObject);
    procedure RzBitBtn1Click(Sender: TObject);
    procedure DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure cxLookupComboBox1PropertiesEditValueChanged(Sender: TObject);
    procedure Ekle1Click(Sender: TObject);
    procedure Edit1Change(Sender: TObject);
    procedure RzBitBtn2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormActivate(Sender: TObject);
    procedure RzBitBtn3Click(Sender: TObject);
    procedure DBGrid2DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
  
  //  procedure DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
    //  DataCol: Integer; Column: TColumn; State: TGridDrawState);
  private
    { Private declarations }
  public
  VAR
  USR,PSW,SRV,KUL,SIF,NOB_TIPI:STRING;
  const sql1='SELECT L.ETIKET,L.TC_KIMLIK_NO FROM DOGUMKAYIT.DA_DRLISTE_NOBET L WHERE '+
              'L.NOBETTE_VAR IS NOT NULL AND L.BRANS=:BRANS AND '+
              'L.TC_KIMLIK_NO NOT IN (select T.TC from DOGUMKAYIT.da_dr_nobet_g t '+
              'WHERE T.TARIH  BETWEEN :TAR1 AND :TAR2 AND T.TC IS NOT NULL union  '+
              'select T.TC from DOGUMKAYIT.da_dr_nobet_g t '+
              'WHERE T.TARIH  BETWEEN :TAR3 AND :TAR4 AND T.TC IS NOT NULL and t.birim=''NB''||:BRANS )'+
             // ' AND L.TC_KIMLIK_NO IN   (select A.TC from DOGUMKAYIT.da_dr_nob_birim_personel A WHERE A.BIRIM=:BIRIM)'+
              ' ORDER BY ETIKET' ;
        sqlcmpz='SELECT L.ETIKET,L.TC_KIMLIK_NO FROM DOGUMKAYIT.DA_DRLISTE_NOBET L WHERE '+
                'L.NOBETTE_VAR IS NOT NULL AND L.BRANS=:BRANS AND '+
                'L.TC_KIMLIK_NO NOT IN ((select T.TC from DOGUMKAYIT.da_dr_nobet_g t '+
                'WHERE T.TARIH  BETWEEN :TAR1 AND :TAR2 AND T.TC IS NOT NULL union  '+
                'select T.TC from DOGUMKAYIT.da_dr_nobet_g t '+
                'WHERE T.TARIH  BETWEEN :TAR3 AND :TAR4 AND T.TC IS NOT NULL and t.birim=''NB''||:BRANS) '+
                'minus select T.TC from DOGUMKAYIT.da_dr_nobet_g t  '+
                'WHERE T.TARIH  BETWEEN :TAR1 AND :TAR2 AND T.TC IS NOT NULL )'+
             //   ' AND L.TC_KIMLIK_NO IN   (select A.TC from DOGUMKAYIT.da_dr_nob_birim_personel A WHERE A.BIRIM=:BIRIM)'+
                'ORDER BY ETIKET';

  end;

var
  FRM_DR_CNOBET: TFRM_DR_CNOBET;


implementation

uses udr_nob_sayi;


{$R *.dfm}



procedure TFRM_DR_CNOBET.ctarihPropertiesChange(Sender: TObject);
begin
tabloloriac;

end;



procedure TFRM_DR_CNOBET.cxLookupComboBox1PropertiesEditValueChanged(
  Sender: TObject);
  VAR
  KOLON,SATIR:INTEGER;
begin
KOLON:=DBGrid1.SelectedIndex;
SATIR:=QTABLO.RecNo;

if (cxLookupComboBox1.EditValue=NULL) OR (DBGrid1.SelectedField.FieldName='TARIH') then
begin
 cxLookupComboBox1.Visible:=false;
  EXIT;
end;

   SQL_INSER.Params[0].Value:=cxLookupComboBox1.EditValue;
   SQL_INSER.Params[1].Value:=QTABLO.FieldByName('TARIH').Value;
   SQL_INSER.Params[2].Value:=DBGrid1.SelectedField.FieldName;
   SQL_INSER.Execute;
   AKGUN.Commit;
   QTABLO.Refresh;
   //cxLookupComboBox1.Visible:=false;
   cxLookupComboBox1.Text:='';


   QTABLO.RecNo:=SATIR;
   DBGrid1.SelectedIndex:=KOLON;
end;







procedure TFRM_DR_CNOBET.DBGrid1DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);

begin



If (gdSelected in state) then // eðer seçili ise
    begin

    DBGrid1.Canvas.Brush.Color := CLRED; // zemin rengi
    DBGrid1.Canvas.Font.Color := clWhite; // font rengi

    Edit1.Text:=QTABLO.FieldByName('TARIH').AsString+DBGrid1.SelectedField.DisplayName;
    FRM_DR_CNOBET.Caption:= DBGRID1.DataSource.DataSet.FieldByName(DBGrid1.SelectedField.DisplayName).AsString  ;

        if RECT.Right+cxLookupComboBox1.Width>DBGrid1.Width then
        cxLookupComboBox1.Left:=RECT.Left-cxLookupComboBox1.Width
    ELSE
        cxLookupComboBox1.Left:=Rect.Right;
        cxLookupComboBox1.Top:=RECT.Bottom+cxLookupComboBox1.Height-5;


    //EDIT2.Text:=DBGrid1.SelectedField.DisplayName;
    end ;

  {  else if (DBGrid1.datasource.dataset.recno mod 2) =0      then
    DBGrid1.Canvas.Brush.Color := clSkyBlue
    else DBGrid1.Canvas.Brush.Color := clWhite; }

     if  DayOfWeek(QTABLO.FieldByName('tarih').AsDateTime)=1 then
       DBGrid1.Canvas.Brush.Color:=clSkyBlue;

     if  DayOfWeek(QTABLO.FieldByName('tarih').AsDateTime)=7 then
       DBGrid1.Canvas.Brush.Color:=clMoneyGreen;




//      if (gdFocused in State) then
//      begin
//        if FRM_DR_CNOBET.Caption= DBGRID1.DataSource.DataSet.FieldByName(DBGrid1.SelectedField.DisplayName).AsString   then
//          DBGrid1.Canvas.Font.Color := clYellow;
//      end;

    DBGrid1.DefaultDrawColumnCell(Rect, DataCol, Column, State);

    TStringGrid(DBGrid1).ScrollBars:=ssBoth;


end;

procedure TFRM_DR_CNOBET.DBGrid2DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin


     if  DayOfWeek(QTABLO_tek.FieldByName('tarih').AsDateTime)=1 then
       DBGrid2.Canvas.Brush.Color:=clSkyBlue;

     if  DayOfWeek(QTABLO_tek.FieldByName('tarih').AsDateTime)=7 then
       DBGrid2.Canvas.Brush.Color:=clMoneyGreen;




//      if (gdFocused in State) then
//      begin
//        if FRM_DR_CNOBET.Caption= DBGRID1.DataSource.DataSet.FieldByName(DBGrid1.SelectedField.DisplayName).AsString   then
//          DBGrid1.Canvas.Font.Color := clYellow;
//      end;

    DBGrid2.DefaultDrawColumnCell(Rect, DataCol, Column, State);

    TStringGrid(DBGrid2).ScrollBars:=ssBoth;



end;

procedure TFRM_DR_CNOBET.Edit1Change(Sender: TObject);
begin

if EBRANSI.EditValue=NULL then  EXIT;
if  DBGrid1.SelectedField.FieldName='NB'+inttostr(EBRANSI.EditValue) THEN
BEGIN

  QDRLOOP.Close;
  QDRLOOP.SQL.Clear;
  QDRLOOP.SQL.Add(sql1);
  QDRLOOP.Params[0].Value:=EBRANSI.EditValue;
  QDRLOOP.Params[1].Value:=QTABLO.FieldByName('TARIH').Value;
  QDRLOOP.Params[2].Value:=QTABLO.FieldByName('TARIH').Value;
  QDRLOOP.Params[3].Value:=QTABLO.FieldByName('TARIH').Value-1;
  QDRLOOP.Params[4].Value:=QTABLO.FieldByName('TARIH').Value+1;
  //QDRLOOP.Params[5].Value:=DBGrid1.SelectedField.FieldName;

END
ELSE if DBGrid1.SelectedField.FieldName='SNOBET' THEN
begin
 /////
end
ELSE if DBGrid1.SelectedField.FieldName='IC'+inttostr(EBRANSI.EditValue) then
BEGIN
  QDRLOOP.Close;
  QDRLOOP.SQL.Clear;
  QDRLOOP.SQL.Add(sqlcmpz);
  QDRLOOP.Params[0].Value:=EBRANSI.EditValue;
  QDRLOOP.Params[1].Value:=QTABLO.FieldByName('TARIH').Value;
  QDRLOOP.Params[2].Value:=QTABLO.FieldByName('TARIH').Value;
  QDRLOOP.Params[3].Value:='01.01.1899';
  QDRLOOP.Params[4].Value:='01.01.1899';
  //QDRLOOP.Params[5].Value:=DBGrid1.SelectedField.FieldName;
END ELSE
BEGIN
  QDRLOOP.Close;
  QDRLOOP.SQL.Clear;
  QDRLOOP.SQL.Add(sql1);
  QDRLOOP.Params[0].Value:=EBRANSI.EditValue;
  QDRLOOP.Params[1].Value:=QTABLO.FieldByName('TARIH').Value;
  QDRLOOP.Params[2].Value:=QTABLO.FieldByName('TARIH').Value;
  QDRLOOP.Params[3].Value:=QTABLO.FieldByName('TARIH').Value-1;
  QDRLOOP.Params[4].Value:=QTABLO.FieldByName('TARIH').Value+0;
  //QDRLOOP.Params[5].Value:=DBGrid1.SelectedField.FieldName;
END;

QDRLOOP.Open;
    if QTABLO.FieldByName('TARIH').AsString<>'' then
      RzLabel3.Caption:=QTABLO.FieldByName('TARIH').AsString+#13+
      FormatDateTime('DDDD', QTABLO.FieldByName('TARIH').Value)
      else
      RzLabel3.Caption:='';

end;

procedure TFRM_DR_CNOBET.Ekle1Click(Sender: TObject);
begin
if DBGrid1.SelectedField.FieldName='TARIH' then  EXIT;
 if NOB_TIPI='O' then   EXIT;
cxLookupComboBox1.Visible:=TRUE;
cxLookupComboBox1.SetFocus;

end;

procedure TFRM_DR_CNOBET.FormActivate(Sender: TObject);
begin
Width:=screen.Width;

end;

procedure TFRM_DR_CNOBET.FormClose(Sender: TObject; var Action: TCloseAction);
begin
AKGUN.Close;
AKGUN.Username:='';
AKGUN.Password:='';
AKGUN.Server:='';
      TRY
      WinExec(pchar('EVRAKTAKIP.exe'+' '+KUL+' '+SIF),SW_SHOWNORMAL);
      EXCEPT
      EXIT;
      END;
     Application.Terminate;
end;

procedure TFRM_DR_CNOBET.FormCreate(Sender: TObject);
var
  i,A: Integer;
begin


 TOP:=0;
LEFT:=0;
ClientHeight:=700;
ClientWidth:=1010;

 if ParamCount<6 then
 BEGIN
  Application.Terminate;
 END;

TRY


USR:=ParamStr(1);
PSW:=ParamStr(2);
SRV:=ParamStr(3);
KUL:=Paramstr(4);
SIF:=Paramstr(5);
NOB_TIPI:=ParamStr(6);
//
//
AKGUN.Username:=USR;
AKGUN.Password:=PSW;
AKGUN.Server:=SRV;



//AKGUN.Username:='DOGAN';
//AKGUN.Password:='19721904';
//AKGUN.Server:='10.42.112.2:1521:ORCL';
//KUL:='SYSDB';


AKGUN.Open;



EXCEPT
Application.Terminate;
END;

QBRANS.Close;
QBRANS.Params[0].Value:=KUL;
qbrans.Open;

QETIKET.Open;

if QBRANS.RecordCount<1 then
begin
{if QBRANS.RecordCount>0 then}  showmessage('Yetkiniz Bulunmamaktadýr !!');
 exit;

end;
EBRANSI.ItemIndex:=0;

TOP:=0;
LEFT:=0;
ClientHeight:=700;
ClientWidth:=1010;
ctarih.Clear;


for i:=1 to 30 do
begin
ctarih.Properties.Items.Add(FormatDateTime('mm.yyyy',(incmonth(date,i-24)))) ;
END;
ctarih.ItemIndex:=23;


end;

procedure TFRM_DR_CNOBET.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
     VAR
  KOLON,SATIR:INTEGER;
begin
KOLON:=DBGrid1.SelectedIndex;
SATIR:=QTABLO.RecNo;

 if NOB_TIPI='O' then   EXIT;
 
if DBGrid1.SelectedField.FieldName='TARIH' then  EXIT;

if (Key = VK_F1 ) then
begin
cxLookupComboBox1.Visible:=TRUE;
cxLookupComboBox1.SetFocus;
end;
if KEY=VK_ESCAPE then
BEGIN
cxLookupComboBox1.Text:='';
cxLookupComboBox1.Visible:=false;
DBGrid1.SetFocus;
END;
if KEY=VK_DELETE then
BEGIN

   SQL_INSER.Params[0].Value:=NULL;
   SQL_INSER.Params[1].Value:=QTABLO.FieldByName('TARIH').Value;
   SQL_INSER.Params[2].Value:=DBGrid1.SelectedField.FieldName;
   SQL_INSER.Execute;
   AKGUN.Commit;
   QTABLO.Refresh;
   QTABLO.RecNo:=SATIR;
   DBGrid1.SelectedIndex:=KOLON;
END;
if key=VK_F10 then
BEGIN

tabloloriac_TEK;
END;



end;

procedure TFRM_DR_CNOBET.RzBitBtn1Click(Sender: TObject);
var
settarih: TDateTime;
begin

if NOB_TIPI='O' then   EXIT;

if (ctarih.Text<='')OR (EBRANSI.EditValue=NULL) then   EXIT;

settarih:=strtodate('01.'+ctarih.Text);
takvim_olustur.Close;



takvim_olustur.Params[0].Value:=settarih;
takvim_olustur.Params[1].Value:=EBRANSI.EditValue;
takvim_olustur.Execute;
AKGUN.Commit;
tabloloriac;

end;

procedure TFRM_DR_CNOBET.RzBitBtn2Click(Sender: TObject);
var
SAYFA: ARRAY OF  TfrxReportPage;
Band: ARRAY OF TfrxBand;
DataBand: ARRAY OF TfrxMasterData;
Memo: ARRAY OF TfrxMemoView;
AB1MEMO,B1MEMO,B2MEMO,B3MEMO,B11MEMO,B13MEMO:TfrxMemoView;
RESIM:TfrxPictureView;
i,X:integer;
SUTUNSAYISI:INTEGER;

begin
SetLength(MEMO,QTABLO.FieldCount*2);
SetLength(BAND,2);
SETLENGTH(SAYFA,2);
SETLENGTH(DataBand,2)   ;

{frxDBDataset1 bileþeni ile tabloya baðlanýyoruz.}
frxDBDataset1.DataSet:=QTABLO;
{ Raporu Temizle}
frxReport1.Clear;
{ FastReportun Veri Aðacýna Tablo Alanlarýný Listele }
frxReport1.DataSets.Add(frxDBDataSet1);
{Rapora Sayfa Ekle}

SAYFA[0] := TfrxReportPage.Create(frxReport1);
{ Baþka Nesnede olmayan bir isim ver}
SAYFA[0].CreateUniqueName;
{Alan ve sayfa geniþliklerini varsayýlan olarak ayarla }
SAYFA[0].SetDefaults;
{Kenar boþluklarýný Ayarla. 10=1 cm}
SAYFA[0].TopMargin:=20;
SAYFA[0].LeftMargin:=20;
{sayfanýn boyutunu ayarlayýn.1 VE 2.PARAMETREYE SIFIR VERÝN.3.PARAMETRE SAYFA GENÝÞLÝÐÝNÝ 4.PARAMETRE SAYFA YÜKSEKLÝÐÝNÝ VERÝR}
SAYFA[0].SetBounds(0,0,29.70,21);
{raporun yatay veya dikey olmasý. kullanabilmek için uses'a Printers ekleyin}
SAYFA[0].Orientation:=poLandscape;//poPortrait;//poLandscape;
{ report title (Sayfa Baþlýðý) bandý Ekle}
SAYFA[0].LeftMargin:=6;
SAYFA[0].RightMargin:=3;
SAYFA[0].TopMargin:=2;
SAYFA[0].BottomMargin:=3;





Band[0] := TfrxReportTitle.Create(SAYFA[0]);
Band[0].CreateUniqueName;
{ band kooordinatlarýný veriyoruz. Top Ve Height özellikleri Yeterli}
Band[0].Top := 0;
Band[0].Height := 130;
resim:=TfrxPictureView.Create(band[0]);
resim.CreateUniqueName;
resim.Picture.LoadFromFile('LOGO.BMP');
resim.SetBounds(0,0,100,100);
Resim.Stretched:=true;

B1MEMO:=TfrxMemoView.Create(BAND[0]);
B1MEMO.CreateUniqueName;
B1MEMO.SetBounds(10,5,800,100);
B1MEMO.HAlign:=haCenter;
B1MEMO.Text:=QETIKET.Fields[0].AsString;// 'T.C. SELÇUKLU KAYMAKAMLIÐI';
B1MEMO.Font.Style:= B1MEMO.Font.Style+[fsBold]   ;
{
B2MEMO:=TfrxMemoView.Create(BAND[0]);
B2MEMO.CreateUniqueName;
B2MEMO.SetBounds(110,50,800,20);
B2MEMO.Text:= 'DR. FARUK SÜKAN DOÐUM VE ÇOCUK HASTANESÝ';
B2MEMO.Font.Style:= B1MEMO.Font.Style+[fsBold] ;
   }
B3MEMO:=TfrxMemoView.Create(BAND[0]);
B3MEMO.CreateUniqueName;
B3MEMO.SetBounds(0,90,70*QTABLO.FieldCount-20,20);
B3MEMO.Text:= 'Aylýk Çalýþma Takvimi';
B3MEMO.HAlign:=haCenter;
B3MEMO.Frame.Typ:=[ftLeft,ftRight,ftTop,ftBottom];

{ Rapor Baþlýðýna Metin Nesnesi Ekle ve Baþlýðý Yaz }
if QTABLO.FieldCount>16 then
begin
SUTUNSAYISI:=16
  end
  else
  begin
    SUTUNSAYISI:=QTABlO.FieldCount-1  ;

end;

for I := 0 to SUTUNSAYISI{QTABLO.FieldCount-1} do
 begin
      Memo[i] := TfrxMemoView.Create(Band[0]);
      Memo[i].CreateUniqueName;
      Memo[i].Text :=QTABLO.Fields[i].FieldName;
      if i=0 then
      Memo[i].SetBounds(i*70,110,50,20)
      else
      Memo[i].SetBounds(i*70-20,110,70,20);

      Memo[i].Font.Size:=9;
      Memo[i].Font.Name:='Arial Narrow';
      Memo[i].Font.Style:=Memo[i].Font.Style+[fsBold] ;
      Memo[i].Frame.Typ:=[FTLEFT,ftRight,ftTop,ftBottom];
      MEMO[i].Color:=clGradientInactiveCaption;//clRed;
      memo[i].HAlign:=haCenter;

      { Metin Nesnesini band geniliðince Geniþlet}
     //Memo.Align := baWidth;

end;

{ masterdata bandý ekle }
DataBand[0] := TfrxMasterData.Create(SAYFA[0]);
DataBand[0].CreateUniqueName;
{masterdata bandýnýn DataSetine forma eklediðimiz frxDBDataSet bileþeninni göster}
DataBand[0].DataSet := frxDBDataSet1;
DataBand[0].Top := 0;
DataBand[0].Height := 20;
{ masterdata ya nesne Ekle}



for I := 0 to SUTUNSAYISI{QTABLO.FieldCount-1} do
begin
    Memo[i+QTABLO.FieldCount] := TfrxMemoView.Create(DataBand[0]);
    Memo[i+QTABLO.FieldCount].CreateUniqueName;
    { Dataya Baðlan }
    Memo[i+QTABLO.FieldCount].DataSet := frxDBDataSet1;
    Memo[i+QTABLO.FieldCount].DataField :=QTABLO.Fields[i].FieldName;
    if i=0 then
    Memo[i+QTABLO.FieldCount].SetBounds( i*70,0, 50, 20)
    else
    Memo[i+QTABLO.FieldCount].SetBounds( i*70-20,0, 70, 20);

    memo[i+QTABLO.FieldCount].Font.Size:=7;
    memo[i+QTABLO.FieldCount].Font.Name:='Arial Narrow';
    Memo[i+QTABLO.FieldCount].Font.Style:=Memo[i+QTABLO.FieldCount].Font.Style+[fsBold];
    Memo[i+QTABLO.FieldCount].Frame.Typ:=[FTLEFT,ftRight,ftTop,ftBottom];
    {nESNEYÝ sAÐA HÝZALA }
    Memo[i+QTABLO.FieldCount].HAlign := haLEFT;
//DayOfWeek(strtodate(<frxDBDataset1."TARIH">))
    memo[i+QTABLO.FieldCount].Highlight.Color:=clMoneyGreen;
    memo[i+QTABLO.FieldCount].Highlight.Font.Size:=7;
    memo[i+QTABLO.FieldCount].Highlight.Font.Color:=clBlack;
    memo[i+QTABLO.FieldCount].Highlight.Font.Name:='Arial Narrow';
    memo[i+QTABLO.FieldCount].Highlight.Font.Style:=Memo[i+QTABLO.FieldCount].Font.Style+[fsBold];
  //  memo[i+QTABLO.FieldCount].Highlight.Condition:=' <frxDBDataset1."TARIH">='+quotedstr('01.01.2012');
    memo[i+QTABLO.FieldCount].Highlight.Condition:='inttostr(DayOfWeek(strtodate(<frxDBDataset1."TARIH">))) in ['+quotedstr('1')+','+quotedstr('7')+']';





end;



band[1]:=TfrxPageFooter.Create(SAYFA[0]);
band[1].CreateUniqueName;
band[1].Top:=0;
band[1].Height:=20;
AB1MEMO:=TfrxMemoView.Create(band[1]);
AB1MEMO.CreateUniqueName;
AB1MEMO.Font.Size:=6;
AB1MEMO.Font.Name:='Arial Narrow';
AB1MEMO.Text:='                                  Fr.257                                                                                                       '+
'Revizyon No: 00                                                                                                          '+
'  Yayýn Tarihi :';
AB1MEMO.Height:=20;
AB1MEMO.Width:=800;
AB1MEMO.HAlign:=haCenter;
    { RAPORU GÖSTER}

    {Rapora 2.Sayfa Ekle}
 //////////////////////////////////////////////////////////////////////////////////////
 ///
 if QTABLO.FieldCount>16 then
 begin


 SUTUNSAYISI:=QTABLO.FieldCount-1;

 SAYFA[1] := TfrxReportPage.Create(frxReport1);

SAYFA[1].CreateUniqueName;
SAYFA[1].SetDefaults;
SAYFA[1].TopMargin:=20;
SAYFA[1].LeftMargin:=20;
SAYFA[1].SetBounds(0,0,29.70,21);
SAYFA[1].Orientation:=poLandscape;
SAYFA[1].LeftMargin:=6;
SAYFA[1].RightMargin:=3;
SAYFA[1].TopMargin:=2;
SAYFA[1].BottomMargin:=3;


Band[4] := TfrxReportTitle.Create(SAYFA[1]);
Band[4].CreateUniqueName;
Band[4].Top := 0;
Band[4].Height := 130;

resim:=TfrxPictureView.Create(Band[4]);
resim.CreateUniqueName;
resim.Picture.LoadFromFile('LOGO.BMP');
resim.SetBounds(0,0,100,100);
Resim.Stretched:=true;

B11MEMO:=TfrxMemoView.Create(Band[4]);
B11MEMO.CreateUniqueName;
B11MEMO.SetBounds(10,5,800,100);
B11MEMO.HAlign:=haCenter;
B11MEMO.Text:=QETIKET.Fields[0].AsString;
B11MEMO.Font.Style:= B11MEMO.Font.Style+[fsBold]   ;

B13MEMO:=TfrxMemoView.Create(BAND[4]);
B13MEMO.CreateUniqueName;
B13MEMO.SetBounds(0,90,70*QTABLO.FieldCount-20,20);
B13MEMO.Text:= 'Aylýk Çalýþma Takvimi';
B13MEMO.HAlign:=haCenter;
B13MEMO.Frame.Typ:=[ftLeft,ftRight,ftTop,ftBottom];

      Memo[50] := TfrxMemoView.Create(Band[4]);
      Memo[50].CreateUniqueName;
      Memo[50].Text :=QTABLO.Fields[0].FieldName;
      Memo[50].SetBounds(0*70,110,50,20);


      Memo[50].Font.Size:=9;
      Memo[50].Font.Name:='Arial Narrow';
      Memo[50].Font.Style:=Memo[50].Font.Style+[fsBold] ;
      Memo[50].Frame.Typ:=[FTLEFT,ftRight,ftTop,ftBottom];
      MEMO[50].Color:=clGradientInactiveCaption;//clRed;
      memo[50].HAlign:=haCenter;



 for I := 16 to {QTABLO.FieldCount-1} SUTUNSAYISI do
 begin
      Memo[i] := TfrxMemoView.Create(Band[4]);
      Memo[i].CreateUniqueName;
      Memo[i].Text :=QTABLO.Fields[i].FieldName;

      Memo[i].SetBounds((i-15)*70-20,110,70,20);



      Memo[i].Font.Size:=9;
      Memo[i].Font.Name:='Arial Narrow';
      Memo[i].Font.Style:=Memo[i].Font.Style+[fsBold] ;
      Memo[i].Frame.Typ:=[FTLEFT,ftRight,ftTop,ftBottom];
      MEMO[i].Color:=clGradientInactiveCaption;//clRed;
      memo[i].HAlign:=haCenter;


end;


DataBand[1] := TfrxMasterData.Create(SAYFA[1]);
DataBand[1].CreateUniqueName;
DataBand[1].DataSet := frxDBDataSet1;
DataBand[1].Top := 0;
DataBand[1].Height := 20;

    Memo[100] := TfrxMemoView.Create(DataBand[1]);
    Memo[100].CreateUniqueName;
    Memo[100].DataSet := frxDBDataSet1;
    Memo[100].DataField :=QTABLO.Fields[0].FieldName;
    Memo[100].SetBounds( 0*70,0, 50, 20);
    memo[100].Font.Size:=7;
    memo[100].Font.Name:='Arial Narrow';
    Memo[100].Font.Style:=Memo[100].Font.Style+[fsBold];
    Memo[100].Frame.Typ:=[FTLEFT,ftRight,ftTop,ftBottom];
    Memo[100].HAlign := haLEFT;
    memo[100].Highlight.Color:=clMoneyGreen;
    memo[100].Highlight.Font.Size:=7;
    memo[100].Highlight.Font.Color:=clBlack;
    memo[100].Highlight.Font.Name:='Arial Narrow';
    memo[100].Highlight.Font.Style:=Memo[100].Font.Style+[fsBold];
    memo[100].Highlight.Condition:='inttostr(DayOfWeek(strtodate(<frxDBDataset1."TARIH">))) in ['+quotedstr('1')+','+quotedstr('7')+']';



for I := 16 to {QTABLO.FieldCount-1} SUTUNSAYISI do
begin
    Memo[i+QTABLO.FieldCount] := TfrxMemoView.Create(DataBand[1]);
    Memo[i+QTABLO.FieldCount].CreateUniqueName;
    Memo[i+QTABLO.FieldCount].DataSet := frxDBDataSet1;
    Memo[i+QTABLO.FieldCount].DataField :=QTABLO.Fields[i].FieldName;



    Memo[i+QTABLO.FieldCount].SetBounds( (i-15)*70-20,0, 70, 20)  ;

    memo[i+QTABLO.FieldCount].Font.Size:=7;
    memo[i+QTABLO.FieldCount].Font.Name:='Arial Narrow';
    Memo[i+QTABLO.FieldCount].Font.Style:=Memo[i+QTABLO.FieldCount].Font.Style+[fsBold];
    Memo[i+QTABLO.FieldCount].Frame.Typ:=[FTLEFT,ftRight,ftTop,ftBottom];
    Memo[i+QTABLO.FieldCount].HAlign := haLEFT;
    memo[i+QTABLO.FieldCount].Highlight.Color:=clMoneyGreen;
    memo[i+QTABLO.FieldCount].Highlight.Font.Size:=7;
    memo[i+QTABLO.FieldCount].Highlight.Font.Color:=clBlack;
    memo[i+QTABLO.FieldCount].Highlight.Font.Name:='Arial Narrow';
    memo[i+QTABLO.FieldCount].Highlight.Font.Style:=Memo[i+QTABLO.FieldCount].Font.Style+[fsBold];
    memo[i+QTABLO.FieldCount].Highlight.Condition:='inttostr(DayOfWeek(strtodate(<frxDBDataset1."TARIH">))) in ['+quotedstr('1')+','+quotedstr('7')+']';



end;

 

band[1]:=TfrxPageFooter.Create(SAYFA[1]);
band[1].CreateUniqueName;
band[1].Top:=0;
band[1].Height:=20;
AB1MEMO:=TfrxMemoView.Create(band[1]);
AB1MEMO.CreateUniqueName;
AB1MEMO.Font.Size:=6;
AB1MEMO.Font.Name:='Arial Narrow';
AB1MEMO.Text:='                                  Fr.257                                                                                                       '+
'Revizyon No: 00                                                                                                          '+
'  Yayýn Tarihi :';
AB1MEMO.Height:=20;
AB1MEMO.Width:=800;
AB1MEMO.HAlign:=haCenter;
end;

FrxReport1.ShowReport;

end;

procedure TFRM_DR_CNOBET.RzBitBtn3Click(Sender: TObject);
begin
frmIstatistik.QISTATISTIK.Close;
frmIstatistik.QISTATISTIK.Params[0].AsString:=ctarih.EditText;
frmIstatistik.QISTATISTIK.Params[1].AsString:=EBRANSI.EditValue;
frmIstatistik.QISTATISTIK.Open;
if frmIstatistik.QISTATISTIK.RecordCount=0 then  EXIT;
frmIstatistik.ShowModal;


end;

procedure TFRM_DR_CNOBET.tabloloriac;
VAR
i:integer;
begin
if (ctarih.Text='') OR (EBRANSI.EditValue=NULL) then  exit;

sql_getir.Close;
sql_getir.Params[1].Value:=ctarih.Text;
sql_getir.Params[2].Value:=EBRANSI.EditValue;
sql_getir.Execute;



QTABLO.Close;
QTABLO.SQL.Clear;
QTABLO.SQL.Add(sql_getir.Params[0].value);
QTABLO.Open;

TStringGrid(dbgrid1).ScrollBars:=ssHorizontal;

for I := 0 to DBGrid1.FieldCount-1 do
begin
if i=0 then
DBGrid1.Columns[i].Width:=50
else
DBGrid1.Columns[i].Width:=66
end;



end;

procedure TFRM_DR_CNOBET.tabloloriac_TEK;
VAR
i:integer;
begin
if (ctarih.Text='') OR (EBRANSI.EditValue=NULL) or (DBGrid1.Fields[DBGrid1.SelectedIndex].Text='') then  exit;
if DBGrid2.Visible=true then
begin
  QTABLO_TEK.Close;
  DBGrid2.Visible:=false;
end else
begin
 DBGrid2.Top:= DBGrid1.Top;
 DBGrid2.Left:=DBGrid1.Left;
 DBGrid2.Height:=DBGrid1.Height;
 DBGrid2.Width:=DBGrid1.Width;

sql_getir_tek.Close;
sql_getir_tek.Params[1].Value:=ctarih.Text;
sql_getir_tek.Params[2].Value:=EBRANSI.EditValue;
SQL_GETIR_TEK.Params[3].Value:=DBGrid1.Fields[DBGrid1.SelectedIndex].Text ;


sql_getir_tek.Execute;

QTABLO_TEK.Close;
QTABLO_TEK.SQL.Clear;
QTABLO_TEK.SQL.Add(sql_getir_tek.Params[0].value);
QTABLO_TEK.Open;
TStringGrid(dbgrid2).ScrollBars:=ssHorizontal;

for I := 0 to DBGrid2.FieldCount-1 do
begin
if i=0 then
DBGrid2.Columns[i].Width:=60
else
DBGrid2.Columns[i].Width:=66
end;
DBGrid2.Visible:=true;
end;

end;

end.
