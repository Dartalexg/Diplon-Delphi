unit DinamicObecpec;

interface
uses
  Winapi.Windows, Winapi.Messages,  Vcl.Menus, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.DBCtrls, Vcl.StdCtrls,
  VclTee.TeeGDIPlus, VCLTee.TeEngine, Vcl.ExtCtrls, VCLTee.TeeProcs,
  VCLTee.Chart, VCLTee.DBChart, Vcl.Grids, Vcl.DBGrids, VCLTee.Series,
  Vcl.ComCtrls,Excel2000,ComObj;
    procedure DinamicObecpecClik;
        procedure OpenDinamicObecpecClik;


implementation

uses Unit4;
procedure XlsStart;
begin
FXlsApp := CreateOleObject('Excel.Application');
end;
//------------------------------------------------------------------------------ ������� �������� ��������������
procedure DinamicObecpecClik;
  var
  x,i,k:integer;
  a,b,c,d,e,f,g:TLineSeries;
  znac:string;
begin
PultUpav.TabSheet7.TabVisible:=True;
znac:='��������'; //������ �������� � ���� ������ 0:n
PultUpav.Chart1.View3d:=False;
PultUpav.Chart1.ClearChart;
  XlsStart;
  FXlsApp.Visible := false;
  //FXlsApp.WorkBooks.Add('');


  b:=TLineSeries.Create(PultUpav.Chart1);
  a:=TLineSeries.Create(PultUpav.Chart1);
  c:=TLineSeries.Create(PultUpav.Chart1);
  d:=TLineSeries.Create(PultUpav.Chart1);
  e:=TLineSeries.Create(PultUpav.Chart1);
  f:=TLineSeries.Create(PultUpav.Chart1);
  g:=TLineSeries.Create(PultUpav.Chart1);
  PultUpav.Chart1.Legend.Title.Text.Text:='�������';
PultUpav.Chart1.Legend.Title.Font.Size:=12;
PultUpav.Chart1.Title.Text.Text:='�������� ���������';
PultUpav.Chart1.Title.Font.Size:=12;
PultUpav.Chart1.AxesList.Left.Title.Text:='';
PultUpav.Chart1.AxesList.Left.Title.Font.Size:=12;
PultUpav.Chart1.AxesList.Bottom.Title.Text:='����';
PultUpav.Chart1.AxesList.Bottom.Title.Font.Size:=12;
;
begin
with PultUpav.StringGridDinamicObecpec do
  for i:=0 to ColCount-1 do
    Cols[i].Clear;
end;
PultUpav.TabSheet7.TabVisible:=False;   //������� �������� ��������������    ���������
for x := 0 to PultUpav.StringGridDinamicObecpec.RowCount-1 do
//ringGrid1.Cells[0,x]:=IntToStr(x);
PultUpav.StringGridDinamicObecpec.Cells[0,0]:='���';
PultUpav.StringGridDinamicObecpec.Cells[1,0]:='2006';
PultUpav.StringGridDinamicObecpec.Cells[2,0]:='2007';
PultUpav.StringGridDinamicObecpec.Cells[3,0]:='2008';
PultUpav.StringGridDinamicObecpec.Cells[4,0]:='2009';
PultUpav.StringGridDinamicObecpec.Cells[5,0]:='2010';
PultUpav.StringGridDinamicObecpec.Cells[6,0]:='2011';
PultUpav.StringGridDinamicObecpec.Cells[7,0]:='2012';
PultUpav.StringGridDinamicObecpec.Cells[8,0]:='2013';
PultUpav.StringGridDinamicObecpec.Cells[9,0]:='2014';
PultUpav.StringGridDinamicObecpec.Cells[10,0]:='2015';
PultUpav.StringGridDinamicObecpec.Cells[11,0]:='2016';
PultUpav.StringGridDinamicObecpec.Cells[12,0]:='2017';
PultUpav.StringGridDinamicObecpec.Cells[13,0]:='2018';
PultUpav.StringGridDinamicObecpec.Cells[14,0]:='2019';
PultUpav.StringGridDinamicObecpec.Cells[15,0]:='2020';
PultUpav.StringGridDinamicObecpec.Cells[16,0]:='2021';
PultUpav.StringGridDinamicObecpec.Cells[17,0]:='2022';
PultUpav.StringGridDinamicObecpec.Cells[18,0]:='2023';
PultUpav.StringGridDinamicObecpec.Cells[19,0]:='2024';
PultUpav.StringGridDinamicObecpec.Cells[20,0]:='2025';
PultUpav.StringGridDinamicObecpec.Cells[21,0]:='2026';
PultUpav.StringGridDinamicObecpec.Cells[22,0]:='2027';
PultUpav.StringGridDinamicObecpec.Cells[23,0]:='2028';
PultUpav.StringGridDinamicObecpec.Cells[24,0]:='2029';
PultUpav.StringGridDinamicObecpec.Cells[25,0]:='2030';
PultUpav.StringGridDinamicObecpec.Cells[26,0]:='2031';

//-----------------------------------------------------------------------------------------------------------------------------
Begin
if PultUpav.CheckBox2.Checked then       //����� ����
begin
PultUpav.TabSheet7.TabVisible:=True;   //������� �������� ��������������    ���������
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'������\����� ����� ���������.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[7].Activate;

 for I := 1 to 26 do
a.AddXY(2006+i,FXlsApp.Cells[4,2+i]);
PultUpav.Chart1.AddSeries(a);
PultUpav.Chart1.View3d:=False;
FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;
a.Title:='����� ����';

  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'������\������_���_�����.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[6].Activate;
 begin
 for i  := 1 to 6 do
If PultUpav.StringGridDinamicObecpec.cells[1,i]='' then
break;
   begin
for K := 1 to 26 do
PultUpav.StringGridDinamicObecpec.cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[69,2+k]);
PultUpav.StringGridDinamicObecpec.cells[0,i]:=znac;
   end;
 FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;
 end;
end
else
End;
//-----------------------------------------------------------------------------------------------------------------------------
Begin
if PultUpav.CheckBox3.Checked then  //���������� �����������
begin
PultUpav.TabSheet7.TabVisible:=True;   //������� �������� ��������������    ���������
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'������\����� ����� ���������.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[7].Activate;

 for I := 1 to 26 do
b.AddXY(2006+i,FXlsApp.Cells[5,2+i]);
PultUpav.Chart1.AddSeries(b);
PultUpav.Chart1.View3d:=False;
FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;
b.Title:='���������� �����������';

  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'������\������_���_�����.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[6].Activate;
 begin
 for i  := 1 to 6 do
If PultUpav.StringGridDinamicObecpec.cells[1,i]='' then
break;
   begin
for K := 1 to 26 do
PultUpav.StringGridDinamicObecpec.cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[70,2+k]);
PultUpav.StringGridDinamicObecpec.cells[0,i]:=znac;
   end;
 FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;
 end;
end
else
End;
//-----------------------------------------------------------------------------------------------------------------------------
Begin
if PultUpav.CheckBox1.Checked then   //����� ������ �����������
begin
PultUpav.TabSheet7.TabVisible:=True;   //������� �������� ��������������    ���������
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'������\����� ����� ���������.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[7].Activate;
 for I := 1 to 26 do
c.AddXY(2006+i,FXlsApp.Cells[6,2+i]);
PultUpav.Chart1.AddSeries(c);
PultUpav.Chart1.View3d:=False;
FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;
c.Title:='����� ������ �����������';

  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'������\������_���_�����.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[6].Activate;
 begin
 for i  := 1 to 6 do
If PultUpav.StringGridDinamicObecpec.cells[1,i]='' then
break;
   begin
for K := 1 to 26 do
PultUpav.StringGridDinamicObecpec.cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[71,2+k]);
PultUpav.StringGridDinamicObecpec.cells[0,i]:=znac;
   end;
 FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;
 end;
end
else
End;
//-----------------------------------------------------------------------------------------------------------------------------
Begin
if PultUpav.CheckBox4.Checked then //��������
begin
PultUpav.TabSheet7.TabVisible:=True;   //������� �������� ��������������    ���������
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'������\����� ����� ���������.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[7].Activate;

 for I := 1 to 26 do
d.AddXY(2006+i,FXlsApp.Cells[7,2+i]);
PultUpav.Chart1.AddSeries(d);
PultUpav.Chart1.View3d:=False;


FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;
d.Title:='��������';

  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'������\������_���_�����.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[6].Activate;
 begin
 for i  := 1 to 6 do
If PultUpav.StringGridDinamicObecpec.cells[1,i]='' then
break;
   begin
for K := 1 to 26 do
PultUpav.StringGridDinamicObecpec.cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[72,2+k]);
PultUpav.StringGridDinamicObecpec.cells[0,i]:=znac;
   end;
 FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;
 end;
end
else
End;
//-----------------------------------------------------------------------------------------------------------------------------
Begin
if PultUpav.CheckBox5.Checked then //�����������
begin
PultUpav.TabSheet7.TabVisible:=True;   //������� �������� ��������������    ���������
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'������\����� ����� ���������.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[7].Activate;

 for I := 1 to 26 do
e.AddXY(2006+i,FXlsApp.Cells[8,2+i]);
PultUpav.Chart1.AddSeries(e);
PultUpav.Chart1.View3d:=False;


FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;
e.Title:='�����������';

  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'������\������_���_�����.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[6].Activate;
 begin
 for i  := 1 to 6 do
If PultUpav.StringGridDinamicObecpec.cells[1,i]='' then
break;
   begin
for K := 1 to 26 do
PultUpav.StringGridDinamicObecpec.cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[73,2+k]);
PultUpav.StringGridDinamicObecpec.cells[0,i]:=znac;
   end;
 FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;
 end;
end
else
End;
//-----------------------------------------------------------------------------------------------------------------------------
Begin
if PultUpav.CheckBox6.Checked then //��������
begin
PultUpav.TabSheet7.TabVisible:=True;   //������� �������� ��������������    ���������
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'������\����� ����� ���������.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[7].Activate;

 for I := 1 to 26 do
f.AddXY(2006+i,FXlsApp.Cells[9,2+i]);
PultUpav.Chart1.AddSeries(f);
PultUpav.Chart1.View3d:=False;


FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;
f.Title:='��������';

  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'������\������_���_�����.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[6].Activate;
 begin
 for i  := 1 to 6 do
If PultUpav.StringGridDinamicObecpec.cells[1,i]='' then
break;
   begin
for K := 1 to 26 do
PultUpav.StringGridDinamicObecpec.cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[74,2+k]);
PultUpav.StringGridDinamicObecpec.cells[0,i]:=znac;
   end;
 FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;
 end;
end
else
End;
//-----------------------------------------------------------------------------------------------------------------------------
Begin
if PultUpav.CheckBox7.Checked then //���������� ��������
begin
PultUpav.TabSheet7.TabVisible:=True;   //������� �������� ��������������    ���������
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'������\����� ����� ���������.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[7].Activate;


 for I := 1 to 26 do
g.AddXY(2006+i,FXlsApp.Cells[10,2+i]);
PultUpav.Chart1.AddSeries(g);
PultUpav.Chart1.View3d:=False;


FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;
g.Title:='���������� ��������';

  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'������\������_���_�����.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[6].Activate;
 begin
 for i  := 1 to 6 do
If PultUpav.StringGridDinamicObecpec.cells[1,i]='' then
break;
   begin
for K := 1 to 26 do
PultUpav.StringGridDinamicObecpec.cells[k,i]:=FormatFloat('0.######',FXlsApp.Cells[75,2+k]);
PultUpav.StringGridDinamicObecpec.cells[0,i]:=znac;
   end;
 FXlsApp.ActiveWorkbook.Save;
FXlsApp.ActiveWorkbook.Close;
 end;
end
else
End;





end;

 procedure OpenDinamicObecpecClik;//�������� �������
 var i:integer;
 begin
  PultUpav.TabSheet7.TabVisible:=False; //

  PultUpav.CheckBox1.Checked:=False;
  PultUpav.CheckBox2.Checked:=False;
  PultUpav.CheckBox3.Checked:=False;
  PultUpav.CheckBox4.Checked:=False;
  PultUpav.CheckBox5.Checked:=False;
  PultUpav.CheckBox6.Checked:=False;
  PultUpav.CheckBox7.Checked:=False;

  with PultUpav.StringGridDinamicObecpec do
  for i:=0 to ColCount-1 do
    Cols[i].Clear;
end;







end.
