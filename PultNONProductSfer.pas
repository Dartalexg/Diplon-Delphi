unit PultNONProductSfer;

interface
uses
  Winapi.Windows, Winapi.Messages,  Vcl.Menus, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.DBCtrls, Vcl.StdCtrls,
  VclTee.TeeGDIPlus, VCLTee.TeEngine, Vcl.ExtCtrls, VCLTee.TeeProcs,
  VCLTee.Chart, VCLTee.DBChart, Vcl.Grids, Vcl.DBGrids, VCLTee.Series,
  Vcl.ComCtrls,Excel2000,ComObj;

  procedure BoxScriptReadGilaZaCheatClik;
  procedure BoxScriptDoliNSClik;


implementation

uses Unit4;
procedure XlsStart;
begin
FXlsApp := CreateOleObject('Excel.Application');
end;
//------------------------------------------------------------------------------ ����� ������� ������������������ �����
  procedure BoxScriptDoliNSClik;//��������  ���� ��
begin
  XlsStart;
  FXlsApp.Visible := false;
  //FXlsApp.WorkBooks.Add('');
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'������\����� ����� ���������.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[5].Activate;
  FXlsApp.Cells[3,12]:=PultUpav.BoxScriptDoliNS.Text;
  FXlsApp.ActiveWorkbook.Save;
  FXlsApp.ActiveWorkbook.Close;
end;
 procedure BoxScriptReadGilaZaCheatClik; //�������� ����� ����� �� ���� ���������
begin
  XlsStart;
  FXlsApp.Visible := false;
  //FXlsApp.WorkBooks.Add('');
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'������\����� ����� ���������.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[5].Activate;
  FXlsApp.Cells[3,16]:=PultUpav.BoxScriptReadGilaZaCheat.Text;
  FXlsApp.ActiveWorkbook.Save;
  FXlsApp.ActiveWorkbook.Close;
end;
end.

