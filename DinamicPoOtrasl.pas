unit DinamicPoOtrasl;
interface
uses
  Winapi.Windows, Winapi.Messages,  Vcl.Menus, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.DBCtrls, Vcl.StdCtrls,
  VclTee.TeeGDIPlus, VCLTee.TeEngine, Vcl.ExtCtrls, VCLTee.TeeProcs,
  VCLTee.Chart, VCLTee.DBChart, Vcl.Grids, Vcl.DBGrids, VCLTee.Series,
  Vcl.ComCtrls,Excel2000,ComObj;

procedure DBLookupComboBoxDinamicClickk;
procedure ComboBoxDinamicClickk;

implementation

uses Unit4;
//------------------------------------------------------------------------------ ������ ��� ������� � ��������
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
//------------------------------------------------------------------------------ ����� �� ����� ������� ������ �������� �� ��������
procedure DBLookupComboBoxDinamicClickk;
begin
PultUpav.ComboBoxDinamic.Enabled:=True;
PultUpav.TabSheetDinamicChart.TabVisible:=False;
if AnsiCompareText('����� ����',Trim(PultUpav.DBLookupComboBoxDinamic.Text)) = 0 then
begin
PultUpav.ComboBoxDinamic.Items.Clear;
PultUpav.ComboBoxDinamic.Text:='';
PultUpav.ComboBoxDinamic.Items.Add('�������� �� ������ ����, ���.�� (�� �)');
PultUpav.ComboBoxDinamic.Items.Add('������������� �������� ����  (���.���)');
PultUpav.ComboBoxDinamic.Items.Add('���� ����� ��������� �� ���� �������- ����, ���.��.');
PultUpav.ComboBoxDinamic.Items.Add('����� ��������� (���.���) (����)');
PultUpav.ComboBoxDinamic.Items.Add('���� ����� ��������� �������� ���� �� ���� ���������- ����, ���.��.');
end
else
begin
PultUpav.ComboBoxDinamic.Items.Clear;
PultUpav.ComboBoxDinamic.Text:='';
PultUpav.ComboBoxDinamic.Items.Add('�������� ������������������ ����� �� ������ ����, ���.��');
PultUpav.ComboBoxDinamic.Items.Add('������������� �������� ���� - ���� (���.���)');
PultUpav.ComboBoxDinamic.Items.Add('��� (���. ���.)-����');
PultUpav.ComboBoxDinamic.Items.Add('���� ����� ��������� �������� ���� - ����, ���.��.');
PultUpav.ComboBoxDinamic.Items.Add('����� ��������� (���.���) (����)');
PultUpav.ComboBoxDinamic.Items.Add('����� ��������� (���.���) (����)');
PultUpav.ComboBoxDinamic.Items.Add('����� ������� (���. ���)');
end;
end;
//------------------------------------------------------------------------------ ����� �� ����� ���������� ������ �������� �� ��������
procedure ComboBoxDinamicClickk;

var i,x:integer;
s:TLineSeries;
begin
PultUpav.TabSheetDinamicChart.TabVisible:=True;
//if not XlsConnect then
    XlsStart;
  FXlsApp.Visible := false;
  //FXlsApp.WorkBooks.Add('');
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'������\������_���_�����.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[3].Activate;
 s:=TLineSeries.Create(PultUpav.ChartDinamic);
 PultUpav.ChartDinamic.ClearChart;



for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
//ringGrid1.Cells[0,x]:=IntToStr(x);
PultUpav.StringGridDinamic.Cells[0,0]:='���';
PultUpav.StringGridDinamic.Cells[0,1]:='��������';
PultUpav.StringGridDinamic.Cells[1,0]:='2006';
PultUpav.StringGridDinamic.Cells[2,0]:='2007';
PultUpav.StringGridDinamic.Cells[3,0]:='2008';
PultUpav.StringGridDinamic.Cells[4,0]:='2009';
PultUpav.StringGridDinamic.Cells[5,0]:='2010';
PultUpav.StringGridDinamic.Cells[6,0]:='2011';
PultUpav.StringGridDinamic.Cells[7,0]:='2012';
PultUpav.StringGridDinamic.Cells[8,0]:='2013';
PultUpav.StringGridDinamic.Cells[9,0]:='2014';
PultUpav.StringGridDinamic.Cells[10,0]:='2015';
PultUpav.StringGridDinamic.Cells[11,0]:='2016';
PultUpav.StringGridDinamic.Cells[12,0]:='2017';
PultUpav.StringGridDinamic.Cells[13,0]:='2018';
PultUpav.StringGridDinamic.Cells[14,0]:='2019';
PultUpav.StringGridDinamic.Cells[15,0]:='2020';
PultUpav.StringGridDinamic.Cells[16,0]:='2021';
PultUpav.StringGridDinamic.Cells[17,0]:='2022';
PultUpav.StringGridDinamic.Cells[18,0]:='2023';
PultUpav.StringGridDinamic.Cells[19,0]:='2024';
PultUpav.StringGridDinamic.Cells[20,0]:='2025';
PultUpav.StringGridDinamic.Cells[21,0]:='2026';
PultUpav.StringGridDinamic.Cells[22,0]:='2027';
PultUpav.StringGridDinamic.Cells[23,0]:='2028';
PultUpav.StringGridDinamic.Cells[24,0]:='2029';
PultUpav.StringGridDinamic.Cells[25,0]:='2030';
PultUpav.StringGridDinamic.Cells[26,0]:='2031';
;

begin
if AnsiCompareText('����� ����',Trim(PultUpav.DBLookupComboBoxDinamic.Text)) = 0 then
begin//1 ����
if AnsiCompareText('�������� �� ������ ����, ���.�� (�� �)',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//11
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.�� (�� �)';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[7,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[7,5+i]);
PultUpav.ChartDinamic.AddSeries(s);
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('������������� �������� ����  (���.���)',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//12
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[8,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[8,5+i]);
PultUpav.ChartDinamic.AddSeries(s);
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('���� ����� ��������� �� ���� �������- ����, ���.��.',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//13
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.��';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[9,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[9,5+i]);
PultUpav.ChartDinamic.AddSeries(s);
PultUpav.ChartDinamic.View3d:=False;
FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//14
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[10,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[10,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin//15
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.��';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[11,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[11,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
end;end;end;end
else
begin
if AnsiCompareText('���������� ����������',Trim(PultUpav.DBLookupComboBoxDinamic.Text)) = 0 then
begin//2 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//21
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='�������� ������������������ ����� �� ������ ���� - ���������� ����������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.��';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[14,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[14,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//22
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='������������� �������� ���� - ���� - ���������� ����������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[15,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[15,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//23
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='��� - ���� - ���������� ����������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[16,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[16,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//24
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='���� ����� ��������� �������� ���� - ���� - ���������� ����������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.��.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[17,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[17,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//25
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='����� ��������� (����) - ���������� ����������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[18,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[18,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//26
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='����� ��������� (����) - ���������� ����������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[19,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[19,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin//27
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='����� ������� - ���������� ����������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[20,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[20,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
end;end;end;end;end;end
else
begin
if AnsiCompareText('����� �����������',Trim(PultUpav.DBLookupComboBoxDinamic.Text)) = 0 then
begin//3 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//31
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='�������� ������������������ ����� �� ������ ���� - ����� �����������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.��';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[24,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[24,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//32
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='������������� �������� ���� - ���� - ����� �����������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[26,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[26,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//33
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='��� - ���� - ����� �����������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[27,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[27,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//34
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='���� ����� ��������� �������� ���� - ���� - ����� �����������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.��.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[28,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[28,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//35
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='����� ��������� (����) - ����� �����������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[29,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[29,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//36
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='����� ��������� (����) - ����� �����������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[30,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[30,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin//37
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='����� ������� - ����� �����������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[31,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[31,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
end;end;end;end;end;end
else
begin
if AnsiCompareText('��������',Trim(PultUpav.DBLookupComboBoxDinamic.Text)) = 0 then
begin//4 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//41
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='�������� ������������������ ����� �� ������ ���� - ��������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.��';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[35,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[35,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//42
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='������������� �������� ���� - ���� - ��������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[37,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[37,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//43
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='��� - ���� - ��������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[38,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[38,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//44
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='���� ����� ��������� �������� ���� - ���� - ��������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.��.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[39,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[39,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//45
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='����� ��������� (����) - ��������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[40,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[40,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//46
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='����� ��������� (����) - ��������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[41,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[41,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin//47
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='����� ������� - ��������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[42,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[42,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
end;end;end;end;end;end
else
begin
if AnsiCompareText('�����������',Trim(PultUpav.DBLookupComboBoxDinamic.Text)) = 0 then
begin//5 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//51
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='�������� ������������������ ����� �� ������ ���� - �����������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.��';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[46,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[46,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//52
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='������������� �������� ���� - ���� - �����������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[47,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[47,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//53
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='��� - ���� - �����������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[48,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[48,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//54
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='���� ����� ��������� �������� ���� - ���� - �����������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.��.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[49,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[49,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//55
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='����� ��������� (����) - �����������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[50,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[50,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//56
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='����� ��������� (����) - �����������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[51,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[51,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin//57
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='����� ������� - �����������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[52,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[52,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
end;end;end;end;end;end
else
begin
if AnsiCompareText('��������',Trim(PultUpav.DBLookupComboBoxDinamic.Text)) = 0 then
begin//6 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//61
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='�������� ������������������ ����� �� ������ ���� - ��������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.��';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[58,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[58,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//62
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='������������� �������� ���� - ���� - ��������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[59,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[59,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//63
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='��� - ���� - ��������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[60,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[60,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//64
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='���� ����� ��������� �������� ���� - ���� - ��������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.��.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[61,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[61,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//65
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='����� ��������� (����) - ��������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[62,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[62,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//66
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='����� ��������� (����) - ��������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[63,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[63,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin//67
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='����� ������� - ��������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[64,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[64,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
end;end;end;end;end;end
else
begin
begin//7 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//71
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='�������� ������������������ ����� �� ������ ���� - ���������� ��������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.��';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[69,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[69,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//72
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='������������� �������� ���� - ���� - ���������� ��������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[70,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[70,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//73
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='��� - ���� - ���������� ��������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[71,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[71,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//74
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='���� ����� ��������� �������� ���� - ���� - ���������� ��������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.��.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[72,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[72,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//75
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='����� ��������� (����) - ���������� ��������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[73,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[73,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(PultUpav.ComboBoxDinamic.Text)) = 0 then
begin//76
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='����� ��������� (����) - ���������� ��������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[74,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[74,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin//77
PultUpav.ChartDinamic.Legend.Title.Text.Text:=PultUpav.ComboBoxDinamic.Text;
PultUpav.ChartDinamic.Legend.Title.Font.Size:=12;
PultUpav.ChartDinamic.Title.Text.Text:='����� ������� - ���������� ��������';
PultUpav.ChartDinamic.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Left.Title.Text:='���.���.';
PultUpav.ChartDinamic.AxesList.Left.Title.Font.Size:=12;
PultUpav.ChartDinamic.AxesList.Bottom.Title.Text:='����';
PultUpav.ChartDinamic.AxesList.Bottom.Title.Font.Size:=12;

for I := 1 to 26 do
for x := 0 to PultUpav.StringGridDinamic.RowCount-1 do
PultUpav.StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[75,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[75,5+i]);
PultUpav.ChartDinamic.AddSeries(s);;
PultUpav.ChartDinamic.View3d:=False;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end;end;end;end;end;end;end
end;end;end;end;end;end;end;

end;










end.
