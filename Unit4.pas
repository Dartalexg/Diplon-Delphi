unit Unit4;

interface

uses
  Winapi.Windows, Winapi.Messages,  Vcl.Menus, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.DBCtrls, Vcl.StdCtrls,
  VclTee.TeeGDIPlus, VCLTee.TeEngine, Vcl.ExtCtrls, VCLTee.TeeProcs,
  VCLTee.Chart, VCLTee.DBChart, Vcl.Grids, Vcl.DBGrids, VCLTee.Series,
  Vcl.ComCtrls,Excel2000,ComObj;

type
  TForm4 = class(TForm)
    MainMenu1: TMainMenu;
    N11: TMenuItem;
    N21: TMenuItem;
    N12: TMenuItem;
    N22: TMenuItem;
    PageControlOsnova: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    PanelBD: TPanel;
    ButtonDimografia: TButton;
    ButtonDinamic: TButton;
    PageControlDinamic: TPageControl;
    TabSheetDinamicTable: TTabSheet;
    Label2: TLabel;
    ComboBoxDinamic: TComboBox;
    StringGridDinamic: TStringGrid;
    TabSheetDinamicChart: TTabSheet;
    ChartDinamic: TChart;
    Label1: TLabel;
    DBLookupComboBoxDinamic: TDBLookupComboBox;
    TabSheetDimografiaTable: TTabSheet;
    StringGridDimografia: TStringGrid;
    Label3: TLabel;
    ComboBoxDimografia: TComboBox;
    Series1: TFastLineSeries;
    procedure FormCreate(Sender: TObject);
    procedure DBLookupComboBoxDinamicClick(Sender: TObject);
    procedure ComboBoxDinamicClick(Sender: TObject);
    procedure ButtonDinamicClick(Sender: TObject);
    procedure ButtonDimografiaClick(Sender: TObject);
    procedure ComboBoxDimografiaClick(Sender: TObject);
    procedure PageControlDinamicChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
Form4: TForm4;
FXlsApp,sheet: variant;
implementation
uses Unit2;
{$R *.dfm}
//-------------------------------------------------------------------------------------------------------------�������
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
//-------------------------------------------------------------------------------------------------------------���������
//------------------------------------------------------------------------------ ��������� ����������� ����������
procedure TForm4.ButtonDimografiaClick(Sender: TObject);
var i:integer;
begin
TabSheetDimografiaTable.TabVisible:=True;
TabSheetDinamicTable.TabVisible:=False;
TabSheetDinamicChart.TabVisible:=False;
ChartDinamic.ClearChart;
PageControlDinamic.ActivePage:=TabSheetDimografiaTable;

StringGridDimografia.Align:=alCustom;
ComboBoxDimografia.Items.Clear;
ComboBoxDimografia.Text:='';
ComboBoxDimografia.Items.Add('������, ����� ����������� ��������� (��� ���)');
ComboBoxDimografia.Items.Add('�����������, ���. ���.');
ComboBoxDimografia.Items.Add('���������� ������� �����������(��� ���)');

ChartDinamic.ClearChart;
with StringGridDimografia do
  for i:=0 to ColCount-1 do
    Cols[i].Clear;
end;
//------------------------------------------------------------------------------ ��������� ����������� �������� �� ��������
procedure TForm4.ButtonDinamicClick(Sender: TObject);
var i:integer;
begin
TabSheetDinamicChart.TabVisible:=False;
TabSheetDimografiaTable.TabVisible:=False;
TabSheetDinamicTable.TabVisible:=True;
ComboBoxDinamic.Text:='';
ChartDinamic.ClearChart;
PageControlDinamic.ActivePage:=TabSheetDinamicTable;

ChartDinamic.ClearChart;
with StringGridDinamic do
  for i:=0 to ColCount-1 do
    Cols[i].Clear;
end;
//------------------------------------------------------------------------------ ����� �� ����� ���������� ������ ����������
procedure TForm4.ComboBoxDimografiaClick(Sender: TObject);
var i,x:integer;
s:TLineSeries;
begin
 TabSheetDinamicChart.TabVisible:=True;
if not XlsConnect then
    XlsStart;
  FXlsApp.Visible := false;
  //FXlsApp.WorkBooks.Add('');
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'������\��_���_�����.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[15].Activate;
 s:=TLineSeries.Create(ChartDinamic);
 ChartDinamic.ClearChart;

for x := 0 to StringGridDimografia.RowCount-1 do
//ringGrid1.Cells[0,x]:=IntToStr(x);
StringGridDimografia.Cells[0,0]:='���';
StringGridDimografia.Cells[0,1]:='��������';
StringGridDimografia.Cells[1,0]:='2006';
StringGridDimografia.Cells[2,0]:='2007';
StringGridDimografia.Cells[3,0]:='2008';
StringGridDimografia.Cells[4,0]:='2009';
StringGridDimografia.Cells[5,0]:='2010';
StringGridDimografia.Cells[6,0]:='2011';
StringGridDimografia.Cells[7,0]:='2012';
StringGridDimografia.Cells[8,0]:='2013';
StringGridDimografia.Cells[9,0]:='2014';
StringGridDimografia.Cells[10,0]:='2015';
StringGridDimografia.Cells[11,0]:='2016';
StringGridDimografia.Cells[12,0]:='2017';
StringGridDimografia.Cells[13,0]:='2018';
StringGridDimografia.Cells[14,0]:='2019';
StringGridDimografia.Cells[15,0]:='2020';
StringGridDimografia.Cells[16,0]:='2021';
StringGridDimografia.Cells[17,0]:='2022';
StringGridDimografia.Cells[18,0]:='2023';
StringGridDimografia.Cells[19,0]:='2024';
StringGridDimografia.Cells[20,0]:='2025';
StringGridDimografia.Cells[21,0]:='2026';
StringGridDimografia.Cells[22,0]:='2027';
StringGridDimografia.Cells[23,0]:='2028';
StringGridDimografia.Cells[24,0]:='2029';
StringGridDimografia.Cells[25,0]:='2030';
StringGridDimografia.Cells[26,0]:='2031'; ;
begin
if AnsiCompareText('������, ����� ����������� ��������� (��� ���)',Trim(ComboBoxDimografia.Text)) = 0 then
begin//11
for I := 1 to 26 do
for x := 0 to StringGridDimografia.RowCount-1 do
StringGridDimografia.Cells[i,1]:=FXlsApp.Cells[71,3+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[71,3+i]);
ChartDinamic.AddSeries(s);
ChartDinamic.View3d:=False;//���� ���� �����
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('�����������, ���. ���.',Trim(ComboBoxDimografia.Text)) = 0 then
begin//12
for I := 1 to 26 do
for x := 0 to StringGridDimografia.RowCount-1 do
StringGridDimografia.Cells[i,1]:=FXlsApp.Cells[72,3+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[72,3+i]);
ChartDinamic.AddSeries(s);
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('���������� ������� �����������(��� ���)',Trim(ComboBoxDimografia.Text)) = 0 then
begin//13
for I := 1 to 26 do
for x := 0 to StringGridDimografia.RowCount-1 do
StringGridDimografia.Cells[i,1]:=FXlsApp.Cells[73,3+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[73,3+i]);
ChartDinamic.AddSeries(s);
FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end;end;end;end;end;
//------------------------------------------------------------------------------ ����� �� ����� ���������� ������ �������� �� ��������
procedure TForm4.ComboBoxDinamicClick(Sender: TObject);

var i,x:integer;
s:TLineSeries;
begin
TabSheetDinamicChart.TabVisible:=True;
if not XlsConnect then
    XlsStart;
  FXlsApp.Visible := false;
  //FXlsApp.WorkBooks.Add('');
  FXlsApp.WorkBooks.open(ExtractFilePath(Application.ExeName)+'������\������_���_�����.xlsx');
  Sheet := FXlsApp.ActiveWorkBook.Sheets;
  Sheet.item[3].Activate;
 s:=TLineSeries.Create(ChartDinamic);
 ChartDinamic.ClearChart;



for x := 0 to StringGridDinamic.RowCount-1 do
//ringGrid1.Cells[0,x]:=IntToStr(x);
StringGridDinamic.Cells[0,0]:='���';
StringGridDinamic.Cells[0,1]:='��������';
StringGridDinamic.Cells[1,0]:='2006';
StringGridDinamic.Cells[2,0]:='2007';
StringGridDinamic.Cells[3,0]:='2008';
StringGridDinamic.Cells[4,0]:='2009';
StringGridDinamic.Cells[5,0]:='2010';
StringGridDinamic.Cells[6,0]:='2011';
StringGridDinamic.Cells[7,0]:='2012';
StringGridDinamic.Cells[8,0]:='2013';
StringGridDinamic.Cells[9,0]:='2014';
StringGridDinamic.Cells[10,0]:='2015';
StringGridDinamic.Cells[11,0]:='2016';
StringGridDinamic.Cells[12,0]:='2017';
StringGridDinamic.Cells[13,0]:='2018';
StringGridDinamic.Cells[14,0]:='2019';
StringGridDinamic.Cells[15,0]:='2020';
StringGridDinamic.Cells[16,0]:='2021';
StringGridDinamic.Cells[17,0]:='2022';
StringGridDinamic.Cells[18,0]:='2023';
StringGridDinamic.Cells[19,0]:='2024';
StringGridDinamic.Cells[20,0]:='2025';
StringGridDinamic.Cells[21,0]:='2026';
StringGridDinamic.Cells[22,0]:='2027';
StringGridDinamic.Cells[23,0]:='2028';
StringGridDinamic.Cells[24,0]:='2029';
StringGridDinamic.Cells[25,0]:='2030';
StringGridDinamic.Cells[26,0]:='2031';
;

begin
if AnsiCompareText('����� ����',Trim(DBLookupComboBoxDinamic.Text)) = 0 then
begin//1 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.�� (�� �)',Trim(ComboBoxDinamic.Text)) = 0 then
begin//11
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[7,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[7,5+i]);
ChartDinamic.AddSeries(s);
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(ComboBoxDinamic.Text)) = 0 then
begin//12
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[8,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[8,5+i]);
ChartDinamic.AddSeries(s);
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� �� ���� �������- ����, ���.��.',Trim(ComboBoxDinamic.Text)) = 0 then
begin//13
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[9,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[9,5+i]);
ChartDinamic.AddSeries(s);
FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBoxDinamic.Text)) = 0 then
begin//14
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[10,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[10,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin//15
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[11,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[11,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
end;end;end;end
else
begin
if AnsiCompareText('���������� ����������',Trim(DBLookupComboBoxDinamic.Text)) = 0 then
begin//2 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(ComboBoxDinamic.Text)) = 0 then
begin//21
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[14,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[14,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(ComboBoxDinamic.Text)) = 0 then
begin//22
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[15,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[15,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(ComboBoxDinamic.Text)) = 0 then
begin//23
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[16,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[16,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(ComboBoxDinamic.Text)) = 0 then
begin//24
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[17,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[17,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBoxDinamic.Text)) = 0 then
begin//25
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[18,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[18,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBoxDinamic.Text)) = 0 then
begin//26
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[19,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[19,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin//27
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[20,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[20,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
end;end;end;end;end;end
else
begin
if AnsiCompareText('����� �����������',Trim(DBLookupComboBoxDinamic.Text)) = 0 then
begin//3 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(ComboBoxDinamic.Text)) = 0 then
begin//31
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[24,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[24,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(ComboBoxDinamic.Text)) = 0 then
begin//32
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[26,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[26,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(ComboBoxDinamic.Text)) = 0 then
begin//33
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[27,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[27,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(ComboBoxDinamic.Text)) = 0 then
begin//34
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[28,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[28,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBoxDinamic.Text)) = 0 then
begin//35
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[29,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[29,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBoxDinamic.Text)) = 0 then
begin//36
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[30,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[30,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin//37
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[31,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[31,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
end;end;end;end;end;end
else
begin
if AnsiCompareText('��������',Trim(DBLookupComboBoxDinamic.Text)) = 0 then
begin//4 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(ComboBoxDinamic.Text)) = 0 then
begin//41
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[35,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[35,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(ComboBoxDinamic.Text)) = 0 then
begin//42
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[37,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[37,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(ComboBoxDinamic.Text)) = 0 then
begin//43
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[38,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[38,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(ComboBoxDinamic.Text)) = 0 then
begin//44
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[39,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[39,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBoxDinamic.Text)) = 0 then
begin//45
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[40,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[40,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBoxDinamic.Text)) = 0 then
begin//46
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[41,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[41,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin//47
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[42,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[42,5+i]);
ChartDinamic.AddSeries(s);;
end
end;end;end;end;end;end
else
begin
if AnsiCompareText('�����������',Trim(DBLookupComboBoxDinamic.Text)) = 0 then
begin//5 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(ComboBoxDinamic.Text)) = 0 then
begin//51
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[46,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[46,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(ComboBoxDinamic.Text)) = 0 then
begin//52
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[47,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[47,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(ComboBoxDinamic.Text)) = 0 then
begin//53
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[48,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[48,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(ComboBoxDinamic.Text)) = 0 then
begin//54
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[49,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[49,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBoxDinamic.Text)) = 0 then
begin//55
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[50,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[50,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBoxDinamic.Text)) = 0 then
begin//56
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[51,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[51,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin//57
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[52,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[52,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
end;end;end;end;end;end
else
begin
if AnsiCompareText('��������',Trim(DBLookupComboBoxDinamic.Text)) = 0 then
begin//6 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(ComboBoxDinamic.Text)) = 0 then
begin//61
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[58,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[58,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(ComboBoxDinamic.Text)) = 0 then
begin//62
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[59,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[59,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(ComboBoxDinamic.Text)) = 0 then
begin//63
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[60,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[60,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(ComboBoxDinamic.Text)) = 0 then
begin//64
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[61,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[61,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBoxDinamic.Text)) = 0 then
begin//65
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[62,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[62,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBoxDinamic.Text)) = 0 then
begin//66
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[63,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[63,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin//67
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[64,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[64,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
end;end;end;end;end;end
else
begin
begin//7 ����
if AnsiCompareText('�������� ������������������ ����� �� ������ ����, ���.��',Trim(ComboBoxDinamic.Text)) = 0 then
begin//71
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[69,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[69,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('������������� �������� ���� - ���� (���.���)',Trim(ComboBoxDinamic.Text)) = 0 then
begin//72
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[70,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[70,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('��� (���. ���.)-����',Trim(ComboBoxDinamic.Text)) = 0 then
begin//73
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[71,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[71,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('���� ����� ��������� �������� ���� - ����, ���.��.',Trim(ComboBoxDinamic.Text)) = 0 then
begin//74
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[72,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[72,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBoxDinamic.Text)) = 0 then
begin//75
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[73,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[73,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin
if AnsiCompareText('����� ��������� (���.���) (����)',Trim(ComboBoxDinamic.Text)) = 0 then
begin//76
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[74,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[74,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end
else
begin//77
for I := 1 to 26 do
for x := 0 to StringGridDinamic.RowCount-1 do
StringGridDinamic.Cells[i,1]:=FXlsApp.Cells[75,5+i];;
for   I := 1 to 26 do
s.AddXY(2005+i,FXlsApp.Cells[75,5+i]);
ChartDinamic.AddSeries(s);;
 FXlsApp.ActiveWorkbook.Save;
 FXlsApp.ActiveWorkbook.Close;
end;end;end;end;end;end;end
end;end;end;end;end;end;end;

end;
//------------------------------------------------------------------------------ ����� �� ����� ������� ������ �������� �� ��������
procedure TForm4.DBLookupComboBoxDinamicClick(Sender: TObject);
begin
ComboBoxDinamic.Enabled:=True;
TabSheetDinamicChart.TabVisible:=False;
if AnsiCompareText('����� ����',Trim(DBLookupComboBoxDinamic.Text)) = 0 then
begin
ComboBoxDinamic.Items.Clear;
ComboBoxDinamic.Text:='';
ComboBoxDinamic.Items.Add('�������� ������������������ ����� �� ������ ����, ���.�� (�� �)');
ComboBoxDinamic.Items.Add('������������� �������� ���� - ���� (���.���)');
ComboBoxDinamic.Items.Add('���� ����� ��������� �������� ���� �� ���� �������- ����, ���.��.');
ComboBoxDinamic.Items.Add('����� ��������� (���.���) (����)');
ComboBoxDinamic.Items.Add('���� ����� ��������� �������� ���� �� ���� ���������- ����, ���.��.');
end
else
begin
ComboBoxDinamic.Items.Clear;
ComboBoxDinamic.Text:='';
ComboBoxDinamic.Items.Add('�������� ������������������ ����� �� ������ ����, ���.��');
ComboBoxDinamic.Items.Add('������������� �������� ���� - ���� (���.���)');
ComboBoxDinamic.Items.Add('��� (���. ���.)-����');
ComboBoxDinamic.Items.Add('���� ����� ��������� �������� ���� - ����, ���.��.');
ComboBoxDinamic.Items.Add('����� ��������� (���.���) (����)');
ComboBoxDinamic.Items.Add('����� ��������� (���.���) (����)');
ComboBoxDinamic.Items.Add('����� ������� (���. ���)');
end;
end;
//------------------------------------------------------------------------------ �������� �����
procedure TForm4.FormCreate(Sender: TObject);
var h,w:real;
begin
//---------------------- ��������� �����
h:=screen.Height;
w:=screen.Width;
Form4.Height:=screen.Height;
Form4.Width:=screen.Width;
Form4.BorderStyle := bsSingle;//������ ��������� �������� �����
Form4.Align := alCustom;//������ ����������� �����
//----------------------���������� �������� ������
PageControlOsnova.Height:= screen.Height;
PageControlOsnova.Width:=  screen.Width;
//----------------------���������� ������ ��
PanelBD.Height:=PageControlOsnova.Height;
//PageControlDinamic.Height:=PageControlOsnova.Height;
TabSheetDimografiaTable.TabVisible:=False;
TabSheetDinamicTable.TabVisible:=False;
//----------------------���������� �������� �� ��������
PageControlDinamic.Width:=Round(PageControlOsnova.Width-PanelBD.Width);//����
StringGridDinamic.Width:=PageControlDinamic.Width;//������ ����
ComboBoxDinamic.Enabled:=False;//���� ���������� ����
TabSheetDinamicChart.TabVisible:=False;//�������� ��������� ����
//PageControlDinamic.ActivePage:=TabSheetDinamicTable;//���������� ������ �������� � �����������
PageControlDinamic.Visible:=True;//��������� ����������� �������� (����)
//----------------------���������� ����������
StringGridDimografia.Width:=PageControlDinamic.Width;//������ ����
end;
procedure TForm4.PageControlDinamicChange(Sender: TObject);
begin

end;

//------------------------------------------------------------------------------

end.