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
 S:integer;
begin
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
