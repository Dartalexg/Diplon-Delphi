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
Memo1.lines.add('Описание сценария11');
end;
2:
begin
Memo1.Clear;
Memo1.lines.add('Описание сценария12');
end;
3:
begin
Memo1.Clear;
Memo1.lines.add('Описание сценария13');
end;
4:
begin
Memo1.Clear;
Memo1.lines.add('Описание сценария14');
end;
5:
begin
Memo1.Clear;
Memo1.lines.add('Описание сценария15');
end;




end;
end;





procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
PultUpav.Enabled:=true;
Box.ItemIndex:=-1;
Memo1.Clear;
end;

end.





