unit OsvoenieMSB;

interface
uses
  Winapi.Windows, Winapi.Messages,  Vcl.Menus, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.DBCtrls, Vcl.StdCtrls,
  VclTee.TeeGDIPlus, VCLTee.TeEngine, Vcl.ExtCtrls, VCLTee.TeeProcs,
  VCLTee.Chart, VCLTee.DBChart, Vcl.Grids, Vcl.DBGrids, VCLTee.Series,
  Vcl.ComCtrls,Excel2000,ComObj;
    procedure OsvoenieMSBOpen;
    procedure Res;


implementation

uses Unit4;
procedure XlsStart;
begin
FXlsApp := CreateOleObject('Excel.Application');
end;
//----------------------------------------------------------------------------------------------
procedure OsvoenieMSBOpen;
var
i,k:integer;
begin
  XlsStart;
  FXlsApp.Visible := false;
  //FXlsApp.WorkBooks.Add('');
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\модШтакельберг\obrabotka_10_TC_NBD.xls');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[9].Activate;
  PultUpav.StringGrid2.Cells[0,0]:='№ года';

for I := 1 to 20 do
PultUpav.StringGrid2.Cells[i,0]:=inttostr(i);

for I := 1 to 10 do
PultUpav.StringGrid2.Cells[0,i]:=FXlsApp.Cells[7+i,3];

for I := 1 to 10 do
for k := 1 to 20 do
PultUpav.StringGrid2.Cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[7+i,3+k]);
  FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;;
//ShowMessage(FXlsApp.Cells[8,3]);    8 строка 3столб

end;


procedure Res;
var
i,k:integer;
begin
  XlsStart;
  FXlsApp.Visible := false;
  //FXlsApp.WorkBooks.Add('');
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\модШтакельберг\obrabotka_10_TC_NBD.xls');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[8].Activate;
case strtoint(Trim(PultUpav.BoxScriptInvesticFB.Text)) of
1:
begin
PultUpav.StringGrid3.Cells[0,0]:='Доходы КБ';
PultUpav.StringGrid3.Cells[0,1]:='ЗП';
for I := 0 to 1 do
for k := 1 to 20 do
PultUpav.StringGrid3.Cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[9+i,6+k]);
//ShowMessage(FXlsApp.Cells[11,7]);
  FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;;
end;
2:
begin
PultUpav.StringGrid3.Cells[0,0]:='Доходы КБ';
PultUpav.StringGrid3.Cells[0,1]:='ЗП';
for I := 0 to 1 do
for k := 1 to 20 do
PultUpav.StringGrid3.Cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[11+i,6+k]);
//ShowMessage(FXlsApp.Cells[11,7]);
  FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;;
end;
3:
begin
PultUpav.StringGrid3.Cells[0,0]:='Доходы КБ';
PultUpav.StringGrid3.Cells[0,1]:='ЗП';
for I := 0 to 1 do
for k := 1 to 20 do
PultUpav.StringGrid3.Cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[13+i,6+k]);
//ShowMessage(FXlsApp.Cells[11,7]);
  FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;;
end;
4:
begin
PultUpav.StringGrid3.Cells[0,0]:='Доходы КБ';
PultUpav.StringGrid3.Cells[0,1]:='ЗП';
for I := 0 to 1 do
for k := 1 to 20 do
PultUpav.StringGrid3.Cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[15+i,6+k]);
//ShowMessage(FXlsApp.Cells[11,7]);
  FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;;
end;
5:
begin
PultUpav.StringGrid3.Cells[0,0]:='Доходы КБ';
PultUpav.StringGrid3.Cells[0,1]:='ЗП';
for I := 0 to 1 do
for k := 1 to 20 do
PultUpav.StringGrid3.Cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[17+i,6+k]);
  FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;;
end;
end;
end;







end.
