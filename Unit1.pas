unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.DBCtrls, Vcl.StdCtrls,
  VclTee.TeeGDIPlus, VCLTee.TeEngine, Vcl.ExtCtrls, VCLTee.TeeProcs,
  VCLTee.Chart, VCLTee.DBChart, Vcl.Grids, Vcl.DBGrids, VCLTee.Series,
  Vcl.ComCtrls,Excel2000,ComObj;

type
  TForm1 = class(TForm)
    Label1: TLabel;
    DBLookupComboBox1: TDBLookupComboBox;
    PageControl: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    Chart: TChart;
    ComboBox1: TComboBox;
    Label2: TLabel;
    StringGrid1: TStringGrid;
    Button1: TButton;
    Button2: TButton;
    procedure FormCreate(Sender: TObject);

    procedure FormResize(Sender: TObject);
    procedure DBLookupComboBox1Click(Sender: TObject);
    procedure ComboBox1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private

  public

  end;

var
Form1: TForm1;
FXlsApp,sheet: variant;
implementation

{$R *.dfm}
uses Unit2,Unit3;


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



procedure TForm1.Button1Click(Sender: TObject);
//var
//s:TLineSeries;
begin
//s:=TLineSeries.Create(Chart);

//s.AddXY(1,2);
//Chart.AddSeries(s);
 // if not XlsConnect then
    //XlsStart;
  //FXlsApp.Visible := true;
  //FXlsApp.WorkBooks.Add('');
  //FXlsApp.WorkBooks.open('C:\Users\Admin\Desktop\2\Win32\Debug\������\������_���_�����.xlsx');
  //Sheet := FXlsApp.ActiveWorkBook.Sheets;
  //Sheet.item[3].Activate;
 //ShowMessage(FXlsApp.Cells[7,6]);
 //FXlsApp.Quit;
end;







procedure TForm1.Button2Click(Sender: TObject);
begin



 (*
if AnsiCompareText('����� ����',Trim(DBLookupComboBox1.Text)) = 0 then
begin//1 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.�� (�� �)',Trim(ComboBox1.Text)) = 0 then
ShowMessage('11')
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(ComboBox1.Text)) = 0 then
ShowMessage('12')
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� �� ���� �������- ����, ���.��.',Trim(ComboBox1.Text)) = 0 then
ShowMessage('13')
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
ShowMessage('14')
else
ShowMessage('15')
end;end;end;end
else
begin
if AnsiCompareText('���������� ����������',Trim(DBLookupComboBox1.Text)) = 0 then
begin//2 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(ComboBox1.Text)) = 0 then
ShowMessage('21')
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(ComboBox1.Text)) = 0 then
ShowMessage('22')
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(ComboBox1.Text)) = 0 then
ShowMessage('23')
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(ComboBox1.Text)) = 0 then
ShowMessage('24')
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
ShowMessage('25')
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
ShowMessage('26')
else
ShowMessage('27')
end;end;end;end;end;end
else
begin
if AnsiCompareText('����� �����������',Trim(DBLookupComboBox1.Text)) = 0 then
begin//3 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(ComboBox1.Text)) = 0 then
ShowMessage('31')
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(ComboBox1.Text)) = 0 then
ShowMessage('32')
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(ComboBox1.Text)) = 0 then
ShowMessage('33')
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(ComboBox1.Text)) = 0 then
ShowMessage('34')
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
ShowMessage('35')
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
ShowMessage('36')
else
ShowMessage('37')
end;end;end;end;end;end
else
begin
if AnsiCompareText('��������',Trim(DBLookupComboBox1.Text)) = 0 then
begin//4 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(ComboBox1.Text)) = 0 then
ShowMessage('41')
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(ComboBox1.Text)) = 0 then
ShowMessage('42')
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(ComboBox1.Text)) = 0 then
ShowMessage('43')
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(ComboBox1.Text)) = 0 then
ShowMessage('44')
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
ShowMessage('45')
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
ShowMessage('46')
else
ShowMessage('47')
end;end;end;end;end;end
else
begin
if AnsiCompareText('�����������',Trim(DBLookupComboBox1.Text)) = 0 then
begin//5 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(ComboBox1.Text)) = 0 then
ShowMessage('51')
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(ComboBox1.Text)) = 0 then
ShowMessage('52')
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(ComboBox1.Text)) = 0 then
ShowMessage('53')
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(ComboBox1.Text)) = 0 then
ShowMessage('54')
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
ShowMessage('55')
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
ShowMessage('56')
else
ShowMessage('57')
end;end;end;end;end;end
else
begin
if AnsiCompareText('��������',Trim(DBLookupComboBox1.Text)) = 0 then
begin//6 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(ComboBox1.Text)) = 0 then
ShowMessage('61')
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(ComboBox1.Text)) = 0 then
ShowMessage('62')
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(ComboBox1.Text)) = 0 then
ShowMessage('63')
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(ComboBox1.Text)) = 0 then
ShowMessage('64')
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
ShowMessage('65')
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
ShowMessage('66')
else
ShowMessage('67')
end;end;end;end;end;end
else
begin
begin//7 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(ComboBox1.Text)) = 0 then
ShowMessage('71')
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(ComboBox1.Text)) = 0 then
ShowMessage('72')
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(ComboBox1.Text)) = 0 then
ShowMessage('73')
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(ComboBox1.Text)) = 0 then
ShowMessage('74')
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
ShowMessage('75')
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
ShowMessage('76')
else
ShowMessage('77')
end;end;end;end;end;end
end;end;end;end;end;end;
*)end;

procedure TForm1.ComboBox1Click(Sender: TObject);

var i,x:integer;
s:TLineSeries;
begin
if not XlsConnect then
    XlsStart;
  FXlsApp.Visible := false;
  //FXlsApp.WorkBooks.Add('');
  FXlsApp.WorkBooks.open('C:\Users\Admin\Desktop\2\Win32\Debug\������\������_���_�����.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[3].Activate;
 s:=TLineSeries.Create(Chart);
 Chart.ClearChart;


TabSheet2.TabVisible:=True;
for x := 0 to StringGrid1.RowCount-1 do
//ringGrid1.Cells[0,x]:=IntToStr(x);
StringGrid1.Cells[0,0]:='���';
StringGrid1.Cells[0,1]:='��������';
StringGrid1.Cells[1,0]:='2006';
StringGrid1.Cells[2,0]:='2007';
StringGrid1.Cells[3,0]:='2008';
StringGrid1.Cells[4,0]:='2009';
StringGrid1.Cells[5,0]:='2010';
StringGrid1.Cells[6,0]:='2011';
StringGrid1.Cells[7,0]:='2012';
StringGrid1.Cells[8,0]:='2013';
StringGrid1.Cells[9,0]:='2014';
StringGrid1.Cells[10,0]:='2015';
StringGrid1.Cells[11,0]:='2016';
StringGrid1.Cells[12,0]:='2017';
StringGrid1.Cells[13,0]:='2018';
StringGrid1.Cells[14,0]:='2019';
StringGrid1.Cells[15,0]:='2020';
StringGrid1.Cells[16,0]:='2021';
StringGrid1.Cells[17,0]:='2022';
StringGrid1.Cells[18,0]:='2023';
StringGrid1.Cells[19,0]:='2024';
StringGrid1.Cells[20,0]:='2025';
StringGrid1.Cells[21,0]:='2026';
StringGrid1.Cells[22,0]:='2027';
StringGrid1.Cells[23,0]:='2028';
StringGrid1.Cells[24,0]:='2029';
StringGrid1.Cells[25,0]:='2030';
StringGrid1.Cells[26,0]:='2031';
;

begin
if AnsiCompareText('����� ����',Trim(DBLookupComboBox1.Text)) = 0 then
begin//1 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.�� (�� �)',Trim(ComboBox1.Text)) = 0 then
begin//11
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[7,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[7,5+i]);
Chart.AddSeries(s);
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(ComboBox1.Text)) = 0 then
begin//12
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[8,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[8,5+i]);
Chart.AddSeries(s);
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� �� ���� �������- ����, ���.��.',Trim(ComboBox1.Text)) = 0 then
begin//13
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[9,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[9,5+i]);
Chart.AddSeries(s);
FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
begin//14
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[10,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[10,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin//15
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[11,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[11,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
end;end;end;end
else
begin
if AnsiCompareText('���������� ����������',Trim(DBLookupComboBox1.Text)) = 0 then
begin//2 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(ComboBox1.Text)) = 0 then
begin//21
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[14,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[14,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(ComboBox1.Text)) = 0 then
begin//22
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[15,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[15,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(ComboBox1.Text)) = 0 then
begin//23
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[16,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[16,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(ComboBox1.Text)) = 0 then
begin//24
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[17,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[17,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
begin//25
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[18,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[18,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
begin//26
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[19,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[19,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin//27
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[20,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[20,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
end;end;end;end;end;end
else
begin
if AnsiCompareText('����� �����������',Trim(DBLookupComboBox1.Text)) = 0 then
begin//3 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(ComboBox1.Text)) = 0 then
begin//31
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[24,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[24,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(ComboBox1.Text)) = 0 then
begin//32
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[26,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[26,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(ComboBox1.Text)) = 0 then
begin//33
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[27,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[27,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(ComboBox1.Text)) = 0 then
begin//34
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[28,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[28,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
begin//35
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[29,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[29,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
begin//36
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[30,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[30,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin//37
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[31,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[31,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
end;end;end;end;end;end
else
begin
if AnsiCompareText('��������',Trim(DBLookupComboBox1.Text)) = 0 then
begin//4 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(ComboBox1.Text)) = 0 then
begin//41
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[35,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[35,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(ComboBox1.Text)) = 0 then
begin//42
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[37,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[37,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(ComboBox1.Text)) = 0 then
begin//43
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[38,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[38,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(ComboBox1.Text)) = 0 then
begin//44
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[39,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[39,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
begin//45
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[40,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[40,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
begin//46
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[41,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[41,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin//47
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[42,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[42,5+i]);
Chart.AddSeries(s);;
end
end;end;end;end;end;end
else
begin
if AnsiCompareText('�����������',Trim(DBLookupComboBox1.Text)) = 0 then
begin//5 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(ComboBox1.Text)) = 0 then
begin//51
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[46,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[46,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(ComboBox1.Text)) = 0 then
begin//52
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[47,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[47,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(ComboBox1.Text)) = 0 then
begin//53
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[48,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[48,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(ComboBox1.Text)) = 0 then
begin//54
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[49,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[49,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
begin//55
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[50,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[50,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
begin//56
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[51,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[51,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin//57
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[52,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[52,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
end;end;end;end;end;end
else
begin
if AnsiCompareText('��������',Trim(DBLookupComboBox1.Text)) = 0 then
begin//6 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(ComboBox1.Text)) = 0 then
begin//61
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[58,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[58,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(ComboBox1.Text)) = 0 then
begin//62
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[59,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[59,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(ComboBox1.Text)) = 0 then
begin//63
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[60,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[60,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(ComboBox1.Text)) = 0 then
begin//64
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[61,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[61,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
begin//65
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[62,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[62,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
begin//66
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[63,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[63,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin//67
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[64,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[64,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
end;end;end;end;end;end
else
begin
begin//7 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(ComboBox1.Text)) = 0 then
begin//71
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[69,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[69,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(ComboBox1.Text)) = 0 then
begin//72
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[70,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[70,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(ComboBox1.Text)) = 0 then
begin//73
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[71,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[71,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(ComboBox1.Text)) = 0 then
begin//74
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[72,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[72,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
begin//75
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[73,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[73,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBox1.Text)) = 0 then
begin//76
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[74,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[74,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin//77
for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[75,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[75,5+i]);
Chart.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end;end;end;end;end;end;end
end;end;end;end;end;end;end;


(*for I := 1 to 26 do
for x := 0 to StringGrid1.RowCount-1 do
StringGrid1.Cells[i,1]:=FXlsApp.Cells[7,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[7,5+i]);
Chart.AddSeries(s);;   *)











// Form3.Show;
// for i := 1 to 26 do
  //        begin
// Chart1.SeriesList[0].AddXY(i, i);
// Label4.Caption:=DBGrid1.Columns[0].Index;
   //       end;

end;

procedure TForm1.DBLookupComboBox1Click(Sender: TObject);
begin

if AnsiCompareText('����� ����',Trim(DBLookupComboBox1.Text)) = 0 then
begin
ComboBox1.Items.Clear;
ComboBox1.Items.Add('�������� ������������������ ����� �� ������ ����, ���.�� (�� �)');
ComboBox1.Items.Add('������������� �������� ���� - ���� (���.���)');
ComboBox1.Items.Add('���� ����� ��������� �������� ���� �� ���� �������- ����, ���.��.');
ComboBox1.Items.Add('����� ��������� (���.���) (����)');
ComboBox1.Items.Add('���� ����� ��������� �������� ���� �� ���� ���������- ����, ���.��.');
end
else
begin
ComboBox1.Items.Clear;
ComboBox1.Items.Add('�������� ������������������ ����� �� ������ ����, ���.��');
ComboBox1.Items.Add('������������� �������� ���� - ���� (���.���)');
ComboBox1.Items.Add('��� (���. ���.)-����');
ComboBox1.Items.Add('���� ����� ��������� �������� ���� - ����, ���.��.');
ComboBox1.Items.Add('����� ��������� (���.���) (����)');
ComboBox1.Items.Add('����� ��������� (���.���) (����)');
ComboBox1.Items.Add('����� ������� (���. ���)');
end;
end;




procedure TForm1.FormCreate(Sender: TObject);//��� ��������
begin
Form1.Position:= poDesktopCenter;
TabSheet2.TabVisible:=False;
Form1.BorderStyle:=bsSingle; //������ �� ���������� �����
//Form1.Height:=98;

end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
 FXlsApp.Quit;
end;

procedure TForm1.FormResize(Sender: TObject);// ��� ��������� �������� �����
begin
//PageControl.Width:=Form1.Width;
//PageControl.Height:=Form1.Height;
//TabSheet2.Height:=PageControl.Height;
//TabSheet2.Width:=PageControl.Width;
Chart.Width:=PageControl.Width;
Chart.Height:=PageControl.Height;
end;

end.





