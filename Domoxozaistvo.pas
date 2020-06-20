unit Domoxozaistvo;

interface
uses
  Winapi.Windows, Winapi.Messages,  Vcl.Menus, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.DBCtrls, Vcl.StdCtrls,
  VclTee.TeeGDIPlus, VCLTee.TeEngine, Vcl.ExtCtrls, VCLTee.TeeProcs,
  VCLTee.Chart, VCLTee.DBChart, Vcl.Grids, Vcl.DBGrids, VCLTee.Series,
  Vcl.ComCtrls,Excel2000,ComObj;
    procedure DomoxozaistvoOpen;
     procedure ComboBox1Clickk;



implementation

uses Unit4;
procedure XlsStart;
begin
FXlsApp := CreateOleObject('Excel.Application');
end;
//----------------------------------------------------------------------------------------------
procedure DomoxozaistvoOpen;
var i,k:integer;
begin
PultUpav.StringGrid4.Cells[0,0]:='Платные услуги';
PultUpav.StringGrid4.Cells[0,1]:='Недвижимость';
  XlsStart;
  FXlsApp.Visible := false;
    FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\Модель Домохозяйств.xlsm');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[1].Activate;
for I := 0 to 1 do
for k := 1 to 15 do
PultUpav.StringGrid4.Cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[7+i,5+k]);
  FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;;
end;
//-------------------------------------------------------------------------------------------------
 procedure ComboBox1Clickk;
var i,k:integer;
begin
PultUpav.StringGrid5.Cells[0,0]:='Платные услуги';
PultUpav.StringGrid5.Cells[0,1]:='Недвижимость';
PultUpav.StringGrid5.Cells[0,2]:='Прочее';
  XlsStart;
  FXlsApp.Visible := false;
    FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\Модель Домохозяйств.xlsm');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[5].Activate;
case strtoint(Trim(PultUpav.ComboBox1.Text)) of

2016:
begin
for I := 0 to 2 do
for k := 1 to 10 do
PultUpav.StringGrid5.Cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[37+i,2+k]);
FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;;
 end;

2017:
begin
for I := 0 to 2 do
for k := 1 to 10 do
PultUpav.StringGrid5.Cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[68+i,2+k]);
FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;;
 end;

2018:
begin
for I := 0 to 2 do
for k := 1 to 10 do
PultUpav.StringGrid5.Cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[100+i,2+k]);
FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;;
 end;

2019:
begin
for I := 0 to 2 do
for k := 1 to 10 do
PultUpav.StringGrid5.Cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[132+i,2+k]);
FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;;
 end;

 2020:
begin
for I := 0 to 2 do
for k := 1 to 10 do
PultUpav.StringGrid5.Cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[164+i,2+k]);
FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;;
 end;

2021:
begin
for I := 0 to 2 do
for k := 1 to 10 do
PultUpav.StringGrid5.Cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[197+i,2+k]);
FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;;
 end;

2022:
begin
for I := 0 to 2 do
for k := 1 to 10 do
PultUpav.StringGrid5.Cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[229+i,2+k]);
FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;;
 end;

 2023:
begin
for I := 0 to 2 do
for k := 1 to 10 do
PultUpav.StringGrid5.Cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[262+i,2+k]);
FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;;
 end;

2024:
begin
for I := 0 to 2 do
for k := 1 to 10 do
PultUpav.StringGrid5.Cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[295+i,2+k]);
FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;;
 end;

 2025:
begin
for I := 0 to 2 do
for k := 1 to 10 do
PultUpav.StringGrid5.Cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[329+i,2+k]);
FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;;
 end;

2026:
begin
for I := 0 to 2 do
for k := 1 to 10 do
PultUpav.StringGrid5.Cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[362+i,2+k]);
FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;;
 end;

 2027:
begin
for I := 0 to 2 do
for k := 1 to 10 do
PultUpav.StringGrid5.Cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[396+i,2+k]);
FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;;
 end;

 2028:
begin
for I := 0 to 2 do
for k := 1 to 10 do
PultUpav.StringGrid5.Cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[430+i,2+k]);
FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;;
 end;

 2029:
begin
for I := 0 to 2 do
for k := 1 to 10 do
PultUpav.StringGrid5.Cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[464+i,2+k]);
FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;;
 end;

 2030:
begin
for I := 0 to 2 do
for k := 1 to 10 do
PultUpav.StringGrid5.Cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[497+i,2+k]);
FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;;
 end;

 end;

end;



end.
