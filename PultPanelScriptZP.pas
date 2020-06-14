unit PultPanelScriptZP;

interface
uses
  Winapi.Windows, Winapi.Messages,  Vcl.Menus, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.DBCtrls, Vcl.StdCtrls,
  VclTee.TeeGDIPlus, VCLTee.TeEngine, Vcl.ExtCtrls, VCLTee.TeeProcs,
  VCLTee.Chart, VCLTee.DBChart, Vcl.Grids, Vcl.DBGrids, VCLTee.Series,
  Vcl.ComCtrls,Excel2000,ComObj;

    procedure BoxHkolaScriptPoctZPClik;
     procedure BoxObheeObrozScriptPoctZPClik;
     procedure BoxBolnicScriptPoctZPClik;
   procedure  BoxPoliclinScriptPoctZPClik;
   procedure  BoxKyltScriptPoctZPClik;
   procedure BoxFizKeltScriptPoctZPClik;
implementation

uses Unit4;
procedure XlsStart;
begin
FXlsApp := CreateOleObject('Excel.Application');
end;
  //------------------------------------------------------------------------------ Пульт вкладка Сценарий роста зарплат в БС

  procedure BoxHkolaScriptPoctZPClik;//Дошкольные учреждения
begin
  XlsStart;
  FXlsApp.Visible := false;
  //FXlsApp.WorkBooks.Add('');
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\Модель_Соц_сфера.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[1].Activate;
  FXlsApp.Cells[29,12]:=PultUpav.BoxHkolaScriptPoctZP.Text;
  FXlsApp.ActiveWorkbook.Save;
  FXlsApp.ActiveWorkbook.Close;
end;
procedure BoxObheeObrozScriptPoctZPClik;// Общее образование   Больницы
begin
  XlsStart;
  FXlsApp.Visible := false;
  //FXlsApp.WorkBooks.Add('');
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\Модель_Соц_сфера.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[1].Activate;
  FXlsApp.Cells[30,12]:=PultUpav.BoxObheeObrozScriptPoctZP.Text;
  FXlsApp.ActiveWorkbook.Save;
  FXlsApp.ActiveWorkbook.Close;
end;
procedure BoxBolnicScriptPoctZPClik;//Больницы
begin
  XlsStart;
  FXlsApp.Visible := false;
  //FXlsApp.WorkBooks.Add('');
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\Модель_Соц_сфера.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[1].Activate;
  FXlsApp.Cells[31,12]:=PultUpav.BoxBolnicScriptPoctZP.Text;
  FXlsApp.ActiveWorkbook.Save;
  FXlsApp.ActiveWorkbook.Close;
end;
 procedure  BoxPoliclinScriptPoctZPClik;//  Полиники
 begin
  XlsStart;
  FXlsApp.Visible := false;
  //FXlsApp.WorkBooks.Add('');
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\Модель_Соц_сфера.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[1].Activate;
  FXlsApp.Cells[32,12]:=PultUpav.BoxPoliclinScriptPoctZP.Text;
  FXlsApp.ActiveWorkbook.Save;
  FXlsApp.ActiveWorkbook.Close;
end;
procedure  BoxKyltScriptPoctZPClik; //Культура
 begin
  XlsStart;
  FXlsApp.Visible := false;
  //FXlsApp.WorkBooks.Add('');
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\Модель_Соц_сфера.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[1].Activate;
  FXlsApp.Cells[33,12]:=PultUpav.BoxKyltScriptPoctZP.Text;
  FXlsApp.ActiveWorkbook.Save;
  FXlsApp.ActiveWorkbook.Close;
end;
procedure BoxFizKeltScriptPoctZPClik; //Физическая культура
 begin
  XlsStart;
  FXlsApp.Visible := false;
  //FXlsApp.WorkBooks.Add('');
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\Модель_Соц_сфера.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[1].Activate;
  FXlsApp.Cells[34,12]:=PultUpav.BoxFizKeltScriptPoctZP.Text;
  FXlsApp.ActiveWorkbook.Save;
  FXlsApp.ActiveWorkbook.Close;
end;
end.

