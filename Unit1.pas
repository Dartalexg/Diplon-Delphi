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
    Panel1: TPanel;
    Label4: TLabel;
    Label5: TLabel;
    Box: TComboBox;
    Memo1: TMemo;

    procedure BoxClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
  private

  public

  end;

var
Form1: TForm1;
FXlsApp,sheet: variant;
implementation

{$R *.dfm}
uses  Unit4;






procedure TForm1.BoxClick(Sender: TObject);
begin
case strtoint(Trim(Box.Text)) of
1:
begin
Memo1.Clear;
Memo1.lines.add('����������� � ������ ������ �� � ����������� ������� �� ���������� ����������������� � ���������� ���������� � ������������� ����������� ��������� ������������ �������������� � ���������������� ������������ ���������.');
end;
2:
begin
Memo1.Clear;
Memo1.lines.add('�������������� ������ (� ��� ������ �������) �� ������������ ������� �� ���������������� ������� ��� �������� ���, ����������� ������� ������������� ������� (������������� ������ ������������).');
Memo1.lines.add('������� ������ ���������� ����������� ������ �� ���� �������������� ���������� ��������� �������, ����������� �� ���������� ��������� �������� ���. ���������� ������������ ������� ������������');
Memo1.lines.add('�� �������� ���������� � ��������� ��������������, �� ������ ����������� ����������� ������������� ������� �������� ���, �� � ����������� �������������� ����������������� ������� � ��������������� ������� ������������ ����������� ������� ��.');

end;
3:
begin
Memo1.Clear;
Memo1.lines.add('���������� �� ������� �������� ���������������� � �������������� ���������� ����� ������������� ������ ���������� ������, ���������� �������������� ������� 10 ���.');
end;
4:
begin
Memo1.Clear;
Memo1.lines.add('���������� �� ������� �������� ���������������� � �������������� ���������� ����� ������������� ������ ���������� ������, ����� �����������, �� ��������� � 3 ��������� ������ ����������� �� �� (� ������� 7 ���)');
end;
5:
begin
Memo1.Clear;
Memo1.lines.add('���������� �� ������� �������� ���������������� � �������������� ���������� ����� ������������� ������ ���������� ������, ����� �����������, �� ��������� � 3 ��������� ������ ����������� �� �� (� ������� 4 ���) ');
end;




end;
end;





procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
PultUpav.Enabled:=true;
Box.ItemIndex:=-1;
Memo1.Clear;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
Memo1.Font.Size:=12;
Memo1.Height:=PultUpav.Memo1.Lines.Count*20;
end;

end.





