unit PultScriptINF;

interface
uses
  Winapi.Windows, Winapi.Messages,  Vcl.Menus, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.DBCtrls, Vcl.StdCtrls,
  VclTee.TeeGDIPlus, VCLTee.TeEngine, Vcl.ExtCtrls, VCLTee.TeeProcs,
  VCLTee.Chart, VCLTee.DBChart, Vcl.Grids, Vcl.DBGrids, VCLTee.Series,
  Vcl.ComCtrls,Excel2000,ComObj;

   procedure BoxScriptINFLClik;
  procedure BitBtn1Clic;
implementation

uses Unit4;
procedure XlsStart;
begin
FXlsApp := CreateOleObject('Excel.Application');
end;
 //------------------------------------------------------------------------------ ����� ������� �������� �������� ���� ��������
procedure BoxScriptINFLClik;
begin
  XlsStart;
  FXlsApp.Visible := false;
  //FXlsApp.WorkBooks.Add('');
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'������\����� ����� ���������.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[5].Activate;
  FXlsApp.Cells[3,4]:=PultUpav.BoxScriptINFL.Text;
  FXlsApp.ActiveWorkbook.Save;
  FXlsApp.ActiveWorkbook.Close;
end;
procedure BitBtn1Clic; //���������
Begin
MessageBox(0, 'Hello '+#13#10+'World1', '�������� ������1 ', mb_IconInformation + mb_OK + mb_TaskModal);
End;
end.
