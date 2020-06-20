//������ � ������� ��������� ������� ������ � ��������� ���������
unit FunctionModel;
interface
uses
  Winapi.Windows, Winapi.Messages,  Vcl.Menus, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.DBCtrls, Vcl.StdCtrls,
  VclTee.TeeGDIPlus, VCLTee.TeEngine, Vcl.ExtCtrls, VCLTee.TeeProcs,
  VCLTee.Chart, VCLTee.DBChart, Vcl.Grids, Vcl.DBGrids, VCLTee.Series,
  Vcl.ComCtrls,Excel2000,ComObj;

procedure DimografiaButtonActiv;
procedure DinamicButtonActiv;
procedure SettingCreate;
implementation

uses Unit4;
//------------------------------------------------------------------------------ ��������� ����������� ����������
procedure DimografiaButtonActiv;
var i:integer;
begin
PultUpav.TabSheetDimografiaTable.TabVisible:=True;
PultUpav.TabSheetDinamicTable.TabVisible:=False;
PultUpav.TabSheetDinamicChart.TabVisible:=False;
PultUpav.ChartDinamic.ClearChart;
PultUpav.PageControlDinamic.ActivePage:=PultUpav.TabSheetDimografiaTable;

PultUpav.StringGridDimografia.Align:=alCustom;
PultUpav.ComboBoxDimografia.Items.Clear;
PultUpav.ComboBoxDimografia.Text:='';
PultUpav.ComboBoxDimografia.Items.Add('����� ����������� ��������� (��� ���)');
PultUpav.ComboBoxDimografia.Items.Add('�����������, ���. ���.');
PultUpav.ComboBoxDimografia.Items.Add('���������� ������� �����������(��� ���)');

PultUpav.ChartDinamic.ClearChart;
with PultUpav.StringGridDimografia do
  for i:=0 to ColCount-1 do
    Cols[i].Clear;
end;
//------------------------------------------------------------------------------ ��������� ����������� �������� �� ��������
procedure DinamicButtonActiv;
var i:integer;
begin
PultUpav.TabSheetDinamicChart.TabVisible:=False;
PultUpav.TabSheetDimografiaTable.TabVisible:=False;
PultUpav.TabSheetDinamicTable.TabVisible:=True;
PultUpav.ComboBoxDinamic.Text:='';
PultUpav.ChartDinamic.ClearChart;
PultUpav.PageControlDinamic.ActivePage:=PultUpav.TabSheetDinamicTable;

PultUpav.ChartDinamic.ClearChart;
with PultUpav.StringGridDinamic do
  for i:=0 to ColCount-1 do
    Cols[i].Clear;
end;
//------------------------------------------------------------------------------ �������� ����� (��������� ��� ��������)
procedure SettingCreate;
var h,w:real;
 k,S:integer;
begin
PultUpav.Memo1.Height:=PultUpav.Memo1.Lines.Count*16;
PultUpav.Memo2.Height:=PultUpav.Memo2.Lines.Count*16;
PultUpav.Memo3.Height:=PultUpav.Memo3.Lines.Count*16;
PultUpav.Memo4.Height:=PultUpav.Memo4.Lines.Count*16;
PultUpav.Memo5.Height:=PultUpav.Memo5.Lines.Count*16;
PultUpav.Memo6.Height:=PultUpav.Memo6.Lines.Count*16;
PultUpav.Memo7.Height:=PultUpav.Memo7.Lines.Count*16;
PultUpav.Memo8.Height:=PultUpav.Memo8.Lines.Count*16;

PultUpav.Label35.Caption:='��� ������ �������� ��� '+PultUpav.BoxYearStartProject.Text;
PultUpav.Image1.Picture.LoadFromFile(ExtractFilePath(Application.ExeName)+'Img\1.JPG');
PultUpav.Image2.Picture.LoadFromFile(ExtractFilePath(Application.ExeName)+'Img\2.JPG');
PultUpav.StringGrid2.ColWidths[0] := 100;
PultUpav.StringGrid4.ColWidths[0] := 100;
PultUpav.StringGrid5.ColWidths[0] := 100;
begin
if PultUpav.TabSheet5.Width<1920 then
begin
PultUpav.ScrollBox6.Width:=1920;
PultUpav.Panel5.Width:=960;
PultUpav.Panel6.Width:=960 ;
end
else
begin
PultUpav.Panel5.Width:=Round(screen.Width/2);
PultUpav.Panel6.Width:=Round(screen.Width/2);
end;
end;




//----------------------���������� ������� ������� ������ � ��
//PultUpav.StringGridTariff.Width:=PultUpav.ScrollBox2.Width;
PultUpav.ScrollBox2.Width:=PultUpav.TabSheet13.Width;
PultUpav.StringGridTariff.ColWidths[0] := 150;
PultUpav.StringGridPayMoney.ColWidths[0] := 150;
//---------------------- ��������� �����
h:=screen.Height;
w:=screen.Width;
PultUpav.Height:=screen.Height;
PultUpav.Width:=screen.Width;
PultUpav.TabSheet7.TabVisible:=False;//������� �������� ��������������    ���������   ����������
PultUpav.PageControlOsnova.TabIndex:=0; //�������� ������ �������
PultUpav.BorderStyle := bsSingle;//������ ��������� �������� �����
PultUpav.Align := alCustom;//������ ����������� �����
//----------------------���������� �������� ������
PultUpav.PageControlOsnova.Height:= screen.Height;
PultUpav.PageControlOsnova.Width:=  screen.Width;
//----------------------���������� ������ ��
PultUpav.PanelBD.Height:=PultUpav.PageControlOsnova.Height;
//PageControlDinamic.Height:=PageControlOsnova.Height;
PultUpav.TabSheetDimografiaTable.TabVisible:=False;
PultUpav.TabSheetDinamicTable.TabVisible:=False;
//----------------------���������� �������� �� ��������
PultUpav.PageControlDinamic.Width:=Round(PultUpav.PageControlOsnova.Width-PultUpav.PanelBD.Width);//����
PultUpav.StringGridDinamic.Width:=PultUpav.PageControlDinamic.Width;//������ ����
PultUpav.ComboBoxDinamic.Enabled:=False;//���� ���������� ����
PultUpav.TabSheetDinamicChart.TabVisible:=False;//�������� ��������� ����
//PageControlDinamic.ActivePage:=TabSheetDinamicTable;//���������� ������ �������� � �����������
PultUpav.PageControlDinamic.Visible:=True;//��������� ����������� �������� (����)
//----------------------���������� ����������
PultUpav.StringGridDimografia.Width:=PultUpav.PageControlDinamic.Width;//������ ����
//----------------------���������� �������� ��������������
PultUpav.StringGridDinamicObecpec.Width:=PultUpav.PageControl1.Width;
//----------------------���������� ������
PultUpav.ScrollBox1.Width:=PultUpav.TabSheet3.Width;
s:=Round(PultUpav.TabSheet3.Width/5);
if PultUpav.TabSheet3.Width<1500 then
else
begin
PultUpav.PultPanelScriptINFL.Width:=s;
PultUpav.PultPanelBudgetRegion.Width:=s;
PultUpav.PultPanelNONProductSfer.Width:=s;
PultUpav.PultPanelScriptRostTarif.Width:=s;
PultUpav.PultPanelScriptZP.Width:=s;
end;
end;

//------------------------------------------------------------------------------
end.
