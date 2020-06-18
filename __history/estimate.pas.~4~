unit estimate;

interface
uses
  Winapi.Windows, Winapi.Messages,  Vcl.Menus, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.DBCtrls, Vcl.StdCtrls,
  VclTee.TeeGDIPlus, VCLTee.TeEngine, Vcl.ExtCtrls, VCLTee.TeeProcs,
  VCLTee.Chart, VCLTee.DBChart, Vcl.Grids, Vcl.DBGrids, VCLTee.Series,
  Vcl.ComCtrls,Excel2000,ComObj;
    procedure estimateopen;


implementation

uses Unit4;
procedure XlsStart;
begin
FXlsApp := CreateOleObject('Excel.Application');
end;
 //------------------------------------------------------------------------------ оценка
procedure estimateopen;
begin
  XlsStart;
  FXlsApp.Visible := false;
    FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\Модель_Соц_сфера.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[6].Activate;

PultUpav.StringGrid1.ColWidths[0] := 200;
PultUpav.StringGrid1.ColWidths[1] := 130;
PultUpav.StringGrid1.ColWidths[2] := 130;
PultUpav.StringGrid1.ColWidths[3] := 200;

PultUpav.StringGrid1.Cells[0,0]:='Наименование отрасли';
PultUpav.StringGrid1.Cells[0,1]:='Жилье';
PultUpav.StringGrid1.Cells[0,2]:='Дошкольные учреждения';
PultUpav.StringGrid1.Cells[0,3]:='Общее образование';
PultUpav.StringGrid1.Cells[0,4]:='Больницы';
PultUpav.StringGrid1.Cells[0,5]:='Поликлиники';
PultUpav.StringGrid1.Cells[0,6]:='Культура';
PultUpav.StringGrid1.Cells[0,7]:='Физическая культура';

PultUpav.StringGrid1.Cells[1,0]:='Эталон';
PultUpav.StringGrid1.Cells[1,1]:='28';
PultUpav.StringGrid1.Cells[1,2]:='0,8';
PultUpav.StringGrid1.Cells[1,3]:='1';
PultUpav.StringGrid1.Cells[1,4]:='0,03';
PultUpav.StringGrid1.Cells[1,5]:='0,05';
PultUpav.StringGrid1.Cells[1,6]:='0,3';
PultUpav.StringGrid1.Cells[1,7]:='0,3';

PultUpav.StringGrid1.Cells[2,0]:='Единицы измерения';
PultUpav.StringGrid1.Cells[2,1]:='кв.м.';
PultUpav.StringGrid1.Cells[2,2]:='мест';
PultUpav.StringGrid1.Cells[2,3]:='мест';
PultUpav.StringGrid1.Cells[2,4]:='мест(коек)';
PultUpav.StringGrid1.Cells[2,5]:='посещений в смену';
PultUpav.StringGrid1.Cells[2,6]:='кв.м.';
PultUpav.StringGrid1.Cells[2,7]:='кв.м.';

PultUpav.StringGrid1.Cells[3,0]:='Прогнозное значение';
PultUpav.StringGrid1.Cells[3,1]:=FormatFloat('0.######',FXlsApp.Cells[69,28]);
PultUpav.StringGrid1.Cells[3,2]:=FormatFloat('0.######',FXlsApp.Cells[70,28]);
PultUpav.StringGrid1.Cells[3,3]:=FormatFloat('0.######',FXlsApp.Cells[71,28]);
PultUpav.StringGrid1.Cells[3,4]:=FormatFloat('0.######',FXlsApp.Cells[72,28]);
PultUpav.StringGrid1.Cells[3,5]:=FormatFloat('0.######',FXlsApp.Cells[73,28]);
PultUpav.StringGrid1.Cells[3,6]:=FormatFloat('0.######',FXlsApp.Cells[74,28]);
PultUpav.StringGrid1.Cells[3,7]:=FormatFloat('0.######',FXlsApp.Cells[75,28]);

FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;

end;

end.
