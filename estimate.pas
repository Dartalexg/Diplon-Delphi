unit estimate;

interface
uses
  Winapi.Windows, Winapi.Messages,  Vcl.Menus, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.DBCtrls, Vcl.StdCtrls,
  VclTee.TeeGDIPlus, VCLTee.TeEngine, Vcl.ExtCtrls, VCLTee.TeeProcs,
  VCLTee.Chart, VCLTee.DBChart, Vcl.Grids, Vcl.DBGrids, VCLTee.Series,
  Vcl.ComCtrls,Excel2000,ComObj;
    procedure estimateopen;
    procedure ComboBoxEstimateClickk;


implementation

uses Unit4;
procedure XlsStart;
begin
FXlsApp := CreateOleObject('Excel.Application');
end;
 //------------------------------------------------------------------------------ оценка
procedure estimateopen;   //при старте программы
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
//------------------------------------------------------------------------------ Выбор из бокса год пункта оценка
procedure ComboBoxEstimateClickk;
begin
case strtoint(Trim(PultUpav.ComboBoxEstimate.Text)) of
2030://2030
begin
  XlsStart;
  FXlsApp.Visible := false;
    FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\Модель_Соц_сфера.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[6].Activate;

PultUpav.StringGrid1.Cells[3,1]:=FormatFloat('0.######',FXlsApp.Cells[69,28]);
PultUpav.StringGrid1.Cells[3,2]:=FormatFloat('0.######',FXlsApp.Cells[70,28]);
PultUpav.StringGrid1.Cells[3,3]:=FormatFloat('0.######',FXlsApp.Cells[71,28]);
PultUpav.StringGrid1.Cells[3,4]:=FormatFloat('0.######',FXlsApp.Cells[72,28]);
PultUpav.StringGrid1.Cells[3,5]:=FormatFloat('0.######',FXlsApp.Cells[73,28]);
PultUpav.StringGrid1.Cells[3,6]:=FormatFloat('0.######',FXlsApp.Cells[74,28]);
PultUpav.StringGrid1.Cells[3,7]:=FormatFloat('0.######',FXlsApp.Cells[75,28]);

FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;

PultUpav.LabelEstimate.Caption:='Год достижения значений '+Trim(PultUpav.ComboBoxEstimate.Text);
end;

2029://2029
begin
  XlsStart;
  FXlsApp.Visible := false;
    FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\Модель_Соц_сфера.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[6].Activate;

PultUpav.StringGrid1.Cells[3,1]:=FormatFloat('0.######',FXlsApp.Cells[69,27]);
PultUpav.StringGrid1.Cells[3,2]:=FormatFloat('0.######',FXlsApp.Cells[70,27]);
PultUpav.StringGrid1.Cells[3,3]:=FormatFloat('0.######',FXlsApp.Cells[71,27]);
PultUpav.StringGrid1.Cells[3,4]:=FormatFloat('0.######',FXlsApp.Cells[72,27]);
PultUpav.StringGrid1.Cells[3,5]:=FormatFloat('0.######',FXlsApp.Cells[73,27]);
PultUpav.StringGrid1.Cells[3,6]:=FormatFloat('0.######',FXlsApp.Cells[74,27]);
PultUpav.StringGrid1.Cells[3,7]:=FormatFloat('0.######',FXlsApp.Cells[75,27]);

FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;

PultUpav.LabelEstimate.Caption:='Год достижения значений '+Trim(PultUpav.ComboBoxEstimate.Text);
end;

2028://2028
begin
  XlsStart;
  FXlsApp.Visible := false;
    FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\Модель_Соц_сфера.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[6].Activate;

PultUpav.StringGrid1.Cells[3,1]:=FormatFloat('0.######',FXlsApp.Cells[69,26]);
PultUpav.StringGrid1.Cells[3,2]:=FormatFloat('0.######',FXlsApp.Cells[70,26]);
PultUpav.StringGrid1.Cells[3,3]:=FormatFloat('0.######',FXlsApp.Cells[71,26]);
PultUpav.StringGrid1.Cells[3,4]:=FormatFloat('0.######',FXlsApp.Cells[72,26]);
PultUpav.StringGrid1.Cells[3,5]:=FormatFloat('0.######',FXlsApp.Cells[73,26]);
PultUpav.StringGrid1.Cells[3,6]:=FormatFloat('0.######',FXlsApp.Cells[74,26]);
PultUpav.StringGrid1.Cells[3,7]:=FormatFloat('0.######',FXlsApp.Cells[75,26]);

FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;

PultUpav.LabelEstimate.Caption:='Год достижения значений '+Trim(PultUpav.ComboBoxEstimate.Text);
end;

2027://2027
begin
  XlsStart;
  FXlsApp.Visible := false;
    FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\Модель_Соц_сфера.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[6].Activate;

PultUpav.StringGrid1.Cells[3,1]:=FormatFloat('0.######',FXlsApp.Cells[69,25]);
PultUpav.StringGrid1.Cells[3,2]:=FormatFloat('0.######',FXlsApp.Cells[70,25]);
PultUpav.StringGrid1.Cells[3,3]:=FormatFloat('0.######',FXlsApp.Cells[71,25]);
PultUpav.StringGrid1.Cells[3,4]:=FormatFloat('0.######',FXlsApp.Cells[72,25]);
PultUpav.StringGrid1.Cells[3,5]:=FormatFloat('0.######',FXlsApp.Cells[73,25]);
PultUpav.StringGrid1.Cells[3,6]:=FormatFloat('0.######',FXlsApp.Cells[74,25]);
PultUpav.StringGrid1.Cells[3,7]:=FormatFloat('0.######',FXlsApp.Cells[75,25]);

FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;

PultUpav.LabelEstimate.Caption:='Год достижения значений '+Trim(PultUpav.ComboBoxEstimate.Text);
end;

2026://2026
begin
  XlsStart;
  FXlsApp.Visible := false;
    FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\Модель_Соц_сфера.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[6].Activate;

PultUpav.StringGrid1.Cells[3,1]:=FormatFloat('0.######',FXlsApp.Cells[69,24]);
PultUpav.StringGrid1.Cells[3,2]:=FormatFloat('0.######',FXlsApp.Cells[70,24]);
PultUpav.StringGrid1.Cells[3,3]:=FormatFloat('0.######',FXlsApp.Cells[71,24]);
PultUpav.StringGrid1.Cells[3,4]:=FormatFloat('0.######',FXlsApp.Cells[72,24]);
PultUpav.StringGrid1.Cells[3,5]:=FormatFloat('0.######',FXlsApp.Cells[73,24]);
PultUpav.StringGrid1.Cells[3,6]:=FormatFloat('0.######',FXlsApp.Cells[74,24]);
PultUpav.StringGrid1.Cells[3,7]:=FormatFloat('0.######',FXlsApp.Cells[75,24]);

FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;

PultUpav.LabelEstimate.Caption:='Год достижения значений '+Trim(PultUpav.ComboBoxEstimate.Text);
end;

2025://2025
begin
  XlsStart;
  FXlsApp.Visible := false;
    FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\Модель_Соц_сфера.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[6].Activate;

PultUpav.StringGrid1.Cells[3,1]:=FormatFloat('0.######',FXlsApp.Cells[69,23]);
PultUpav.StringGrid1.Cells[3,2]:=FormatFloat('0.######',FXlsApp.Cells[70,23]);
PultUpav.StringGrid1.Cells[3,3]:=FormatFloat('0.######',FXlsApp.Cells[71,23]);
PultUpav.StringGrid1.Cells[3,4]:=FormatFloat('0.######',FXlsApp.Cells[72,23]);
PultUpav.StringGrid1.Cells[3,5]:=FormatFloat('0.######',FXlsApp.Cells[73,23]);
PultUpav.StringGrid1.Cells[3,6]:=FormatFloat('0.######',FXlsApp.Cells[74,23]);
PultUpav.StringGrid1.Cells[3,7]:=FormatFloat('0.######',FXlsApp.Cells[75,23]);

FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;

PultUpav.LabelEstimate.Caption:='Год достижения значений '+Trim(PultUpav.ComboBoxEstimate.Text);
end;

2024://2024
begin
  XlsStart;
  FXlsApp.Visible := false;
    FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\Модель_Соц_сфера.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[6].Activate;

PultUpav.StringGrid1.Cells[3,1]:=FormatFloat('0.######',FXlsApp.Cells[69,22]);
PultUpav.StringGrid1.Cells[3,2]:=FormatFloat('0.######',FXlsApp.Cells[70,22]);
PultUpav.StringGrid1.Cells[3,3]:=FormatFloat('0.######',FXlsApp.Cells[71,22]);
PultUpav.StringGrid1.Cells[3,4]:=FormatFloat('0.######',FXlsApp.Cells[72,22]);
PultUpav.StringGrid1.Cells[3,5]:=FormatFloat('0.######',FXlsApp.Cells[73,22]);
PultUpav.StringGrid1.Cells[3,6]:=FormatFloat('0.######',FXlsApp.Cells[74,22]);
PultUpav.StringGrid1.Cells[3,7]:=FormatFloat('0.######',FXlsApp.Cells[75,22]);

FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;

PultUpav.LabelEstimate.Caption:='Год достижения значений '+Trim(PultUpav.ComboBoxEstimate.Text);
end;

2023://2023
begin
  XlsStart;
  FXlsApp.Visible := false;
    FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\Модель_Соц_сфера.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[6].Activate;

PultUpav.StringGrid1.Cells[3,1]:=FormatFloat('0.######',FXlsApp.Cells[69,21]);
PultUpav.StringGrid1.Cells[3,2]:=FormatFloat('0.######',FXlsApp.Cells[70,21]);
PultUpav.StringGrid1.Cells[3,3]:=FormatFloat('0.######',FXlsApp.Cells[71,21]);
PultUpav.StringGrid1.Cells[3,4]:=FormatFloat('0.######',FXlsApp.Cells[72,21]);
PultUpav.StringGrid1.Cells[3,5]:=FormatFloat('0.######',FXlsApp.Cells[73,21]);
PultUpav.StringGrid1.Cells[3,6]:=FormatFloat('0.######',FXlsApp.Cells[74,21]);
PultUpav.StringGrid1.Cells[3,7]:=FormatFloat('0.######',FXlsApp.Cells[75,21]);

FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;

PultUpav.LabelEstimate.Caption:='Год достижения значений '+Trim(PultUpav.ComboBoxEstimate.Text);
end;

2022://2022
begin
  XlsStart;
  FXlsApp.Visible := false;
    FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\Модель_Соц_сфера.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[6].Activate;

PultUpav.StringGrid1.Cells[3,1]:=FormatFloat('0.######',FXlsApp.Cells[69,20]);
PultUpav.StringGrid1.Cells[3,2]:=FormatFloat('0.######',FXlsApp.Cells[70,20]);
PultUpav.StringGrid1.Cells[3,3]:=FormatFloat('0.######',FXlsApp.Cells[71,20]);
PultUpav.StringGrid1.Cells[3,4]:=FormatFloat('0.######',FXlsApp.Cells[72,20]);
PultUpav.StringGrid1.Cells[3,5]:=FormatFloat('0.######',FXlsApp.Cells[73,20]);
PultUpav.StringGrid1.Cells[3,6]:=FormatFloat('0.######',FXlsApp.Cells[74,20]);
PultUpav.StringGrid1.Cells[3,7]:=FormatFloat('0.######',FXlsApp.Cells[75,20]);

FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;

PultUpav.LabelEstimate.Caption:='Год достижения значений '+Trim(PultUpav.ComboBoxEstimate.Text);
end;

2021://2021
begin
  XlsStart;
  FXlsApp.Visible := false;
    FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\Модель_Соц_сфера.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[6].Activate;

PultUpav.StringGrid1.Cells[3,1]:=FormatFloat('0.######',FXlsApp.Cells[69,19]);
PultUpav.StringGrid1.Cells[3,2]:=FormatFloat('0.######',FXlsApp.Cells[70,19]);
PultUpav.StringGrid1.Cells[3,3]:=FormatFloat('0.######',FXlsApp.Cells[71,19]);
PultUpav.StringGrid1.Cells[3,4]:=FormatFloat('0.######',FXlsApp.Cells[72,19]);
PultUpav.StringGrid1.Cells[3,5]:=FormatFloat('0.######',FXlsApp.Cells[73,19]);
PultUpav.StringGrid1.Cells[3,6]:=FormatFloat('0.######',FXlsApp.Cells[74,19]);
PultUpav.StringGrid1.Cells[3,7]:=FormatFloat('0.######',FXlsApp.Cells[75,19]);

FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;

PultUpav.LabelEstimate.Caption:='Год достижения значений '+Trim(PultUpav.ComboBoxEstimate.Text);
end;

2020://2020
begin
  XlsStart;
  FXlsApp.Visible := false;
    FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\Модель_Соц_сфера.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[6].Activate;

PultUpav.StringGrid1.Cells[3,1]:=FormatFloat('0.######',FXlsApp.Cells[69,18]);
PultUpav.StringGrid1.Cells[3,2]:=FormatFloat('0.######',FXlsApp.Cells[70,18]);
PultUpav.StringGrid1.Cells[3,3]:=FormatFloat('0.######',FXlsApp.Cells[71,18]);
PultUpav.StringGrid1.Cells[3,4]:=FormatFloat('0.######',FXlsApp.Cells[72,18]);
PultUpav.StringGrid1.Cells[3,5]:=FormatFloat('0.######',FXlsApp.Cells[73,18]);
PultUpav.StringGrid1.Cells[3,6]:=FormatFloat('0.######',FXlsApp.Cells[74,18]);
PultUpav.StringGrid1.Cells[3,7]:=FormatFloat('0.######',FXlsApp.Cells[75,18]);

FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;

PultUpav.LabelEstimate.Caption:='Год достижения значений '+Trim(PultUpav.ComboBoxEstimate.Text);
end;


end;
end;



end.
