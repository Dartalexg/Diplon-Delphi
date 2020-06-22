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
Memo1.lines.add('Инерционный – объемы помощи ФБ и собственных доходов от имеющегося производственного и природного потенциала – экстраполяция сложившихся тенденций федерального финансирования и функционирования региональной экономики.');
end;
2:
begin
Memo1.Clear;
Memo1.lines.add('Единовременная помощь (в год старта проекта) из федерального бюджета на инфраструктурные проекты для освоения МСБ, собственные расходы регионального бюджета (двухуровневая модель Штакельберга).');
Memo1.lines.add('Краевой бюджет наращивает собственные доходы за счет дополнительных «сырьевых» бюджетных потоков, поступающих от реализации программы освоения МСБ. Инвестиции федерального бюджета направляются');
Memo1.lines.add('на развитие территории – создается инфраструктура, не только открывающая возможность рентабельного запуска проектов МСБ, но и порождающая дополнительные мультипликативные эффекты и соответствующий прирост традиционных собственных доходов КБ.');

end;
3:
begin
Memo1.Clear;
Memo1.lines.add('Инвестиции ФБ первого сценария трансформируются в дополнительные трансферты сверх традиционного объема финансовой помощи, равномерно распределяемые течение 10 лет.');
end;
4:
begin
Memo1.Clear;
Memo1.lines.add('Инвестиции ФБ первого сценария трансформируются в дополнительные трансферты сверх традиционного объема финансовой помощи, более интенсивный, по сравнению с 3 сценарием график поступлений из ФБ (в течении 7 лет)');
end;
5:
begin
Memo1.Clear;
Memo1.lines.add('Инвестиции ФБ первого сценария трансформируются в дополнительные трансферты сверх традиционного объема финансовой помощи, более интенсивный, по сравнению с 3 сценарием график поступлений из ФБ (в течении 4 лет) ');
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





