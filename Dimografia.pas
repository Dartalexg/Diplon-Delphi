unit Dimografia;

interface
uses
  Winapi.Windows, Winapi.Messages,  Vcl.Menus, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.DBCtrls, Vcl.StdCtrls,
  VclTee.TeeGDIPlus, VCLTee.TeEngine, Vcl.ExtCtrls, VCLTee.TeeProcs,
  VCLTee.Chart, VCLTee.DBChart, Vcl.Grids, Vcl.DBGrids, VCLTee.Series,
  Vcl.ComCtrls,Excel2000,ComObj;

  procedure ComboBoxDimografiaClickk;

implementation



uses Unit4;
function XlsConnect: boolean;
begin
Result := False;
try
FXlsApp := GetActiveOleObject('Excel.Application');
Result := True;
except
end;
end;
procedure XlsStart;
begin
FXlsApp := CreateOleObject('Excel.Application');

end;
//------------------------------------------------------------------------------ ����� �� ����� ���������� ������ ����������
procedure ComboBoxDimografiaClickk;
var i,x:integer;
s:TLineSeries;
begin
 PultUpav.TabSheetDinamicChart.TabVisible:=True;
//if not XlsConnect then ��������
  XlsStart;
  FXlsApp.Visible := false;
  //FXlsApp.WorkBooks.Add('');
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'������\��_���_�����.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[15].Activate;
 s:=TLineSeries.Create(PultUpav.ChartDinamic);
 PultUpav.ChartDinamic.ClearChart;

for x := 0 to PultUpav.StringGridDimografia.RowCount-1 do
//ringGrid1.Cells[0,x]:=IntToStr(x);
PultUpav.StringGridDimografia.Cells[0,0]:='���';
PultUpav.StringGridDimografia.Cells[0,1]:='��������';
PultUpav.StringGridDimografia.Cells[1,0]:='2006';
PultUpav.StringGridDimografia.Cells[2,0]:='2007';
PultUpav.StringGridDimografia.Cells[3,0]:='2008';
PultUpav.StringGridDimografia.Cells[4,0]:='2009';
PultUpav.StringGridDimografia.Cells[5,0]:='2010';
PultUpav.StringGridDimografia.Cells[6,0]:='2011';
PultUpav.StringGridDimografia.Cells[7,0]:='2012';
PultUpav.StringGridDimografia.Cells[8,0]:='2013';
PultUpav.StringGridDimografia.Cells[9,0]:='2014';
PultUpav.StringGridDimografia.Cells[10,0]:='2015';
PultUpav.StringGridDimografia.Cells[11,0]:='2016';
PultUpav.StringGridDimografia.Cells[12,0]:='2017';
PultUpav.StringGridDimografia.Cells[13,0]:='2018';
PultUpav.StringGridDimografia.Cells[14,0]:='2019';
PultUpav.StringGridDimografia.Cells[15,0]:='2020';
PultUpav.StringGridDimografia.Cells[16,0]:='2021';
PultUpav.StringGridDimografia.Cells[17,0]:='2022';
PultUpav.StringGridDimografia.Cells[18,0]:='2023';
PultUpav.StringGridDimografia.Cells[19,0]:='2024';
PultUpav.StringGridDimografia.Cells[20,0]:='2025';
PultUpav.StringGridDimografia.Cells[21,0]:='2026';
PultUpav.StringGridDimografia.Cells[22,0]:='2027';
PultUpav.StringGridDimografia.Cells[23,0]:='2028';
PultUpav.StringGridDimografia.Cells[24,0]:='2029';
PultUpav.StringGridDimografia.Cells[25,0]:='2030';
PultUpav.StringGridDimografia.Cells[26,0]:='2031'; ;
begin
if AnsiCompareText('����� ����������� ��������� (��� ���)',Trim(PultUpav.ComboBoxDimografia.Text)) = 0 then
begin//11
for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDimografia.RowCount-1 do
PultUpav.StringGridDimografia.Cells[i,1]:=FXlsApp.Cells[71,3+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[71,3+i]);
PultUpav.ChartDinamic.AddSeries(s);
PultUpav.ChartDinamic.View3d:=False;//���� ���� �����
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('�����������, ���. ���.',Trim(PultUpav.ComboBoxDimografia.Text)) = 0 then
begin//12
for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDimografia.RowCount-1 do
PultUpav.StringGridDimografia.Cells[i,1]:=FXlsApp.Cells[72,3+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[72,3+i]);
PultUpav.ChartDinamic.AddSeries(s);
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('���������� ������� �����������(��� ���)',Trim(PultUpav.ComboBoxDimografia.Text)) = 0 then
begin//13
for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDimografia.RowCount-1 do
PultUpav.StringGridDimografia.Cells[i,1]:=FXlsApp.Cells[73,3+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[73,3+i]);
PultUpav.ChartDinamic.AddSeries(s);
PultUpav.ChartDinamic.View3d:=False;
FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end;end;end;end;end;






end.
