unit Unit3;

interface

uses
  Winapi.Windows, Winapi.Messages,  Vcl.Menus, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.DBCtrls, Vcl.StdCtrls,
  VclTee.TeeGDIPlus, VCLTee.TeEngine, Vcl.ExtCtrls, VCLTee.TeeProcs,
  VCLTee.Chart, VCLTee.DBChart, Vcl.Grids, Vcl.DBGrids, VCLTee.Series,
  Vcl.ComCtrls,Excel2000,ComObj;

type
  TForm3 = class(TForm)
    Panel1: TPanel;
    Label4: TLabel;
    Label5: TLabel;
    Box: TComboBox;
    Memo1: TMemo;
    procedure FormDestroy(Sender: TObject);
    procedure BoxClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form3: TForm3;
   FXlsApp,sheet: variant;
implementation

{$R *.dfm}

uses Unit4;



procedure TForm3.BoxClick(Sender: TObject);
begin
case strtoint(Trim(Box.Text)) of
1:
begin
Memo1.Clear;
Memo1.lines.add('���������������� �������� - ������� �������� �������, ��������� ���� �� 1-3%');
end;
2:
begin
Memo1.Clear;
Memo1.lines.add('����������� �������� - ������� �������� ��. 4-5%');
end;
3:
begin
Memo1.Clear;
Memo1.lines.add('��������������� �������� - ��������� ���� �� 3-5%');
end;





end;
end;

procedure TForm3.FormClose(Sender: TObject; var Action: TCloseAction);
begin
PultUpav.Enabled:=true;
Box.ItemIndex:=-1;
Memo1.Clear;
end;

procedure TForm3.FormCreate(Sender: TObject);
begin
Box.ItemIndex:=-1;
Memo1.Clear;
Memo1.Font.Size:=12;
Memo1.Height:=PultUpav.Memo1.Lines.Count*20;
end;

procedure TForm3.FormDestroy(Sender: TObject);
begin
PultUpav.Enabled:=true;
end;

end.
