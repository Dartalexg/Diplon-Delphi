unit Tariff;
interface
uses
  Winapi.Windows, Winapi.Messages,  Vcl.Menus, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.DBCtrls, Vcl.StdCtrls,
  VclTee.TeeGDIPlus, VCLTee.TeEngine, Vcl.ExtCtrls, VCLTee.TeeProcs,
  VCLTee.Chart, VCLTee.DBChart, Vcl.Grids, Vcl.DBGrids, VCLTee.Series,
  Vcl.ComCtrls,Excel2000,ComObj;

procedure TariffOpen;


implementation

uses Unit4;
//------------------------------------------------------------------------------ Фукция для конекта с экселями
procedure XlsStart;
begin
FXlsApp := CreateOleObject('Excel.Application');
end;
//------------------------------------------------------------------------------ Обнавление табл диаграммы вкладки тарифов
procedure TariffOpen;
  var
  x,i,k:integer;
  a,b,c,d,e,f,g:TLineSeries;
begin
  XlsStart;
  FXlsApp.Visible := false;
    FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\Модель_Соц_сфера.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[7].Activate;

PultUpav.ChartTariff.ClearChart;
with PultUpav.StringGridTariff do
  for i:=0 to ColCount-1 do
    Cols[i].Clear;

  b:=TLineSeries.Create(PultUpav.ChartTariff);
  a:=TLineSeries.Create(PultUpav.ChartTariff);
  c:=TLineSeries.Create(PultUpav.ChartTariff);
  d:=TLineSeries.Create(PultUpav.ChartTariff);
  e:=TLineSeries.Create(PultUpav.ChartTariff);
  f:=TLineSeries.Create(PultUpav.ChartTariff);
  g:=TLineSeries.Create(PultUpav.ChartTariff);
PultUpav.ChartTariff.Legend.Title.Text.Text:='Легенда';
PultUpav.ChartTariff.Legend.Title.Font.Size:=12;
PultUpav.ChartTariff.Title.Text.Text:='Динамика  тарифа для населения';
PultUpav.ChartTariff.Title.Font.Size:=12;
PultUpav.ChartTariff.AxesList.Left.Title.Text:='';
PultUpav.ChartTariff.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartTariff.AxesList.Bottom.Title.Text:='Года';
PultUpav.ChartTariff.AxesList.Bottom.Title.Font.Size:=12;

for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
//ringGrid1.Cells[0,x]:=IntToStr(x);
PultUpav.StringGridTariff.Cells[0,0]:='Год';
PultUpav.StringGridTariff.Cells[0,1]:='Жилой фонд';
PultUpav.StringGridTariff.Cells[0,2]:='Дошкольные учреждения';
PultUpav.StringGridTariff.Cells[0,3]:='Общее образование';
PultUpav.StringGridTariff.Cells[0,4]:='Больницы';
PultUpav.StringGridTariff.Cells[0,5]:='Поликлиники';
PultUpav.StringGridTariff.Cells[0,6]:='Культура';
PultUpav.StringGridTariff.Cells[0,7]:='Физическая культура';
PultUpav.StringGridTariff.Cells[1,0]:='2006';
PultUpav.StringGridTariff.Cells[2,0]:='2007';
PultUpav.StringGridTariff.Cells[3,0]:='2008';
PultUpav.StringGridTariff.Cells[4,0]:='2009';
PultUpav.StringGridTariff.Cells[5,0]:='2010';
PultUpav.StringGridTariff.Cells[6,0]:='2011';
PultUpav.StringGridTariff.Cells[7,0]:='2012';
PultUpav.StringGridTariff.Cells[8,0]:='2013';
PultUpav.StringGridTariff.Cells[9,0]:='2014';
PultUpav.StringGridTariff.Cells[10,0]:='2015';
PultUpav.StringGridTariff.Cells[11,0]:='2016';
PultUpav.StringGridTariff.Cells[12,0]:='2017';
PultUpav.StringGridTariff.Cells[13,0]:='2018';
PultUpav.StringGridTariff.Cells[14,0]:='2019';
PultUpav.StringGridTariff.Cells[15,0]:='2020';
PultUpav.StringGridTariff.Cells[16,0]:='2021';
PultUpav.StringGridTariff.Cells[17,0]:='2022';
PultUpav.StringGridTariff.Cells[18,0]:='2023';
PultUpav.StringGridTariff.Cells[19,0]:='2024';
PultUpav.StringGridTariff.Cells[20,0]:='2025';
PultUpav.StringGridTariff.Cells[21,0]:='2026';
PultUpav.StringGridTariff.Cells[22,0]:='2027';
PultUpav.StringGridTariff.Cells[23,0]:='2028';
PultUpav.StringGridTariff.Cells[24,0]:='2029';
PultUpav.StringGridTariff.Cells[25,0]:='2030';
PultUpav.StringGridTariff.Cells[26,0]:='2031';

 for I := 1 to 26 do
begin
a.AddXY(2006+i,FXlsApp.Cells[8,6+i]); //Жилой фонд
PultUpav.ChartTariff.AddSeries(a);
PultUpav.ChartTariff.View3d:=False;
a.Title:='Жилой фонд';
PultUpav.StringGridTariff.cells[i,1]:=FormatFloat('0.######',FXlsApp.Cells[8,6+i]);

b.AddXY(2006+i,FXlsApp.Cells[9,6+i]); //Дошкольные учреждения
PultUpav.ChartTariff.AddSeries(b);
PultUpav.ChartTariff.View3d:=False;
b.Title:='Дошкольные учреждения';
PultUpav.StringGridTariff.cells[i,2]:=FormatFloat('0.######',FXlsApp.Cells[9,6+i]);

c.AddXY(2006+i,FXlsApp.Cells[10,6+i]); //Общее образование
PultUpav.ChartTariff.AddSeries(c);
PultUpav.ChartTariff.View3d:=False;
c.Title:='Общее образование';
PultUpav.StringGridTariff.cells[i,3]:=FormatFloat('0.######',FXlsApp.Cells[10,6+i]);

d.AddXY(2006+i,FXlsApp.Cells[11,6+i]); //Больницы
PultUpav.ChartTariff.AddSeries(d);
PultUpav.ChartTariff.View3d:=False;
d.Title:='Больницы';
PultUpav.StringGridTariff.cells[i,4]:=FormatFloat('0.######',FXlsApp.Cells[11,6+i]);

e.AddXY(2006+i,FXlsApp.Cells[12,6+i]); //Поликлиники
PultUpav.ChartTariff.AddSeries(e);
PultUpav.ChartTariff.View3d:=False;
e.Title:='Поликлиники';
PultUpav.StringGridTariff.cells[i,5]:=FormatFloat('0.######',FXlsApp.Cells[12,6+i]);

f.AddXY(2006+i,FXlsApp.Cells[13,6+i]); //Культура
PultUpav.ChartTariff.AddSeries(f);
PultUpav.ChartTariff.View3d:=False;
f.Title:='Культура';
PultUpav.StringGridTariff.cells[i,6]:=FormatFloat('0.######',FXlsApp.Cells[13,6+i]);

g.AddXY(2006+i,FXlsApp.Cells[14,6+i]); //Физическая культура
PultUpav.ChartTariff.AddSeries(g);
PultUpav.ChartTariff.View3d:=False;
g.Title:='Физическая культура';
PultUpav.StringGridTariff.cells[i,7]:=FormatFloat('0.######',FXlsApp.Cells[14,6+i]);

end;
FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;
end;


end.
