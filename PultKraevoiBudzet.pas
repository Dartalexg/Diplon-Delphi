unit PultKraevoiBudzet;

interface
uses
  Winapi.Windows, Winapi.Messages,  Vcl.Menus, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.DBCtrls, Vcl.StdCtrls,
  VclTee.TeeGDIPlus, VCLTee.TeEngine, Vcl.ExtCtrls, VCLTee.TeeProcs,
  VCLTee.Chart, VCLTee.DBChart, Vcl.Grids, Vcl.DBGrids, VCLTee.Series,
  Vcl.ComCtrls,Excel2000,ComObj;
    procedure BoxScriptTransferAndInvestClik;
   procedure BoxYearStartProjectClik;
    procedure BoxScriptInvesticFBClik ;
   procedure BoxTempPoctDoxodOtStartEkonomClik;
  procedure BitBtn2Clic;

implementation

uses Unit4;
procedure XlsStart;
begin
FXlsApp := CreateOleObject('Excel.Application');
end;
 //------------------------------------------------------------------------------ Пульт вкладка Сценарии инфляции бокс сценарии
procedure BoxYearStartProjectClik; //Год старта проектов МСБ
begin
  XlsStart;
  FXlsApp.Visible := false;
  //FXlsApp.WorkBooks.Add('');
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\темпы роста сравнение.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[5].Activate;
  FXlsApp.Cells[4,4]:=PultUpav.BoxYearStartProject.Text;
  FXlsApp.ActiveWorkbook.Save;
  FXlsApp.ActiveWorkbook.Close;
  PultUpav.Label35.Caption:='Год старта проектов МСБ '+PultUpav.BoxYearStartProject.Text;
end;
procedure BoxScriptInvesticFBClik;//Cценарий инвестиций ФБ в МСБ
begin
  XlsStart;
  FXlsApp.Visible := false;
  //FXlsApp.WorkBooks.Add('');
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\темпы роста сравнение.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[5].Activate;
  FXlsApp.Cells[5,4]:=PultUpav.BoxScriptInvesticFB.Text;
  FXlsApp.ActiveWorkbook.Save;
  FXlsApp.ActiveWorkbook.Close;
end;
procedure BoxTempPoctDoxodOtStartEkonomClik;//Темп роста собственных доходов от старой экономики
begin
  XlsStart;
  FXlsApp.Visible := false;
  //FXlsApp.WorkBooks.Add('');
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\темпы роста сравнение.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[5].Activate;
  FXlsApp.Cells[3,7]:=PultUpav.BoxTempPoctDoxodOtStartEkonom.Text;
  FXlsApp.ActiveWorkbook.Save;
  FXlsApp.ActiveWorkbook.Close;
end;
procedure BoxScriptTransferAndInvestClik;//Сценарий трансфертов и инвестиций
begin
  XlsStart;
  FXlsApp.Visible := false;
  //FXlsApp.WorkBooks.Add('');
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'Модель\темпы роста сравнение.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[5].Activate;
  FXlsApp.Cells[6,7]:=PultUpav.BoxScriptTransferAndInvest.Text;
  FXlsApp.ActiveWorkbook.Save;
  FXlsApp.ActiveWorkbook.Close;
end;
 procedure BitBtn2Clic; //Подсказка
 Begin
   MessageBox(0, 'Hello '+#13#10+'World2', 'Название капшин2 ', mb_IconInformation + mb_OK + mb_TaskModal);
 End;
end.

