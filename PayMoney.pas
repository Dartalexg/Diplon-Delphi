unit PayMoney;
interface
uses
  Winapi.Windows, Winapi.Messages,  Vcl.Menus, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.DBCtrls, Vcl.StdCtrls,
  VclTee.TeeGDIPlus, VCLTee.TeEngine, Vcl.ExtCtrls, VCLTee.TeeProcs,
  VCLTee.Chart, VCLTee.DBChart, Vcl.Grids, Vcl.DBGrids, VCLTee.Series,
  Vcl.ComCtrls,Excel2000,ComObj;

procedure PayMoneyOpen;


implementation

uses Unit4;
//------------------------------------------------------------------------------ ������ ��� ������� � ��������
procedure XlsStart;
begin
FXlsApp := CreateOleObject('Excel.Application');
end;
//------------------------------------------------------------------------------ ���������� ���� ��������� ������� ��
procedure PayMoneyOpen;
var
  x,i,k:integer;
  a,b,c,d,e,f,g:TLineSeries;
begin
  XlsStart;
  FXlsApp.Visible := false;
    FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'������\������_���_�����.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[7].Activate;

PultUpav.ChartPayMoney.ClearChart;
with PultUpav.StringGridPayMoney do
  for i:=0 to ColCount-1 do
    Cols[i].Clear;

  b:=TLineSeries.Create(PultUpav.ChartPayMoney);
  a:=TLineSeries.Create(PultUpav.ChartPayMoney);
  c:=TLineSeries.Create(PultUpav.ChartPayMoney);
  d:=TLineSeries.Create(PultUpav.ChartPayMoney);
  e:=TLineSeries.Create(PultUpav.ChartPayMoney);
  f:=TLineSeries.Create(PultUpav.ChartPayMoney);
  g:=TLineSeries.Create(PultUpav.ChartPayMoney);
PultUpav.ChartPayMoney.Legend.Title.Text.Text:='�������';
PultUpav.ChartPayMoney.Legend.Title.Font.Size:=12;
PultUpav.ChartPayMoney.Title.Text.Text:='�������� ��������  � ���������� ��';
PultUpav.ChartPayMoney.Title.Font.Size:=12;
PultUpav.ChartPayMoney.AxesList.Left.Title.Text:='';
PultUpav.ChartPayMoney.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartPayMoney.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartPayMoney.AxesList.Bottom.Title.Font.Size:=12;

for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
//ringGrid1.Cells[0,x]:=IntToStr(x);
PultUpav.StringGridPayMoney.Cells[0,0]:='���';
PultUpav.StringGridPayMoney.Cells[0,1]:='���������� ����������';
PultUpav.StringGridPayMoney.Cells[0,2]:='����� �����������';
PultUpav.StringGridPayMoney.Cells[0,3]:='��������';
PultUpav.StringGridPayMoney.Cells[0,4]:='�����������';
PultUpav.StringGridPayMoney.Cells[0,5]:='��������';
PultUpav.StringGridPayMoney.Cells[0,6]:='���������� ��������';
PultUpav.StringGridPayMoney.Cells[1,0]:='2006';
PultUpav.StringGridPayMoney.Cells[2,0]:='2007';
PultUpav.StringGridPayMoney.Cells[3,0]:='2008';
PultUpav.StringGridPayMoney.Cells[4,0]:='2009';
PultUpav.StringGridPayMoney.Cells[5,0]:='2010';
PultUpav.StringGridPayMoney.Cells[6,0]:='2011';
PultUpav.StringGridPayMoney.Cells[7,0]:='2012';
PultUpav.StringGridPayMoney.Cells[8,0]:='2013';
PultUpav.StringGridPayMoney.Cells[9,0]:='2014';
PultUpav.StringGridPayMoney.Cells[10,0]:='2015';
PultUpav.StringGridPayMoney.Cells[11,0]:='2016';
PultUpav.StringGridPayMoney.Cells[12,0]:='2017';
PultUpav.StringGridPayMoney.Cells[13,0]:='2018';
PultUpav.StringGridPayMoney.Cells[14,0]:='2019';
PultUpav.StringGridPayMoney.Cells[15,0]:='2020';
PultUpav.StringGridPayMoney.Cells[16,0]:='2021';
PultUpav.StringGridPayMoney.Cells[17,0]:='2022';
PultUpav.StringGridPayMoney.Cells[18,0]:='2023';
PultUpav.StringGridPayMoney.Cells[19,0]:='2024';
PultUpav.StringGridPayMoney.Cells[20,0]:='2025';
PultUpav.StringGridPayMoney.Cells[21,0]:='2026';
PultUpav.StringGridPayMoney.Cells[22,0]:='2027';
PultUpav.StringGridPayMoney.Cells[23,0]:='2028';
PultUpav.StringGridPayMoney.Cells[24,0]:='2029';
PultUpav.StringGridPayMoney.Cells[25,0]:='2030';
PultUpav.StringGridPayMoney.Cells[26,0]:='2031';

for I := 1 to 26 do
begin

b.AddXY(2006+i,FXlsApp.Cells[22,6+i]); //���������� ����������
PultUpav.ChartPayMoney.AddSeries(b);
PultUpav.ChartPayMoney.View3d:=False;
b.Title:='���������� ����������';
PultUpav.StringGridPayMoney.cells[i,1]:=FormatFloat('0.######',FXlsApp.Cells[22,6+i]);

c.AddXY(2006+i,FXlsApp.Cells[23,6+i]); //����� �����������
PultUpav.ChartPayMoney.AddSeries(c);
PultUpav.ChartPayMoney.View3d:=False;
c.Title:='����� �����������';
PultUpav.StringGridPayMoney.cells[i,2]:=FormatFloat('0.######',FXlsApp.Cells[23,6+i]);

d.AddXY(2006+i,FXlsApp.Cells[24,6+i]); //��������
PultUpav.ChartPayMoney.AddSeries(d);
PultUpav.ChartPayMoney.View3d:=False;
d.Title:='��������';
PultUpav.StringGridPayMoney.cells[i,3]:=FormatFloat('0.######',FXlsApp.Cells[24,6+i]);

e.AddXY(2006+i,FXlsApp.Cells[25,6+i]); //�����������
PultUpav.ChartPayMoney.AddSeries(e);
PultUpav.ChartPayMoney.View3d:=False;
e.Title:='�����������';
PultUpav.StringGridPayMoney.cells[i,4]:=FormatFloat('0.######',FXlsApp.Cells[25,6+i]);

f.AddXY(2006+i,FXlsApp.Cells[26,6+i]); //��������
PultUpav.ChartPayMoney.AddSeries(f);
PultUpav.ChartPayMoney.View3d:=False;
f.Title:='��������';
PultUpav.StringGridPayMoney.cells[i,5]:=FormatFloat('0.######',FXlsApp.Cells[26,6+i]);

g.AddXY(2006+i,FXlsApp.Cells[27,6+i]); //���������� ��������
PultUpav.ChartPayMoney.AddSeries(g);
PultUpav.ChartPayMoney.View3d:=False;
g.Title:='���������� ��������';
PultUpav.StringGridPayMoney.cells[i,6]:=FormatFloat('0.######',FXlsApp.Cells[27,6+i]);

end;
FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;
end;

end.
