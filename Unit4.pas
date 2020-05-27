unit Unit4;

interface

uses
  Winapi.Windows, Winapi.Messages,  Vcl.Menus, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.DBCtrls, Vcl.StdCtrls,
  VclTee.TeeGDIPlus, VCLTee.TeEngine, Vcl.ExtCtrls, VCLTee.TeeProcs,
  VCLTee.Chart, VCLTee.DBChart, Vcl.Grids, Vcl.DBGrids, VCLTee.Series,
  Vcl.ComCtrls,Excel2000,ComObj;

type
  TPultUpav = class(TForm)
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
    procedure FormDestroy(Sender: TObject);




  private
    { Private declarations }
  public
  end;

var
PultUpav: TPultUpav;
FXlsApp,sheet: variant;
implementation
uses Unit2, FunctionModel, DinamicPoOtrasl, Dimografia;
{$R *.dfm}
 //------------------------------------------------------------------------------ ������ ��� ������� � ��������

//------------------------------------------------------------------------------ ��������� ����������� ����������
procedure TPultUpav.ButtonDimografiaClick(Sender: TObject);
begin
DimografiaButtonActiv;// FunctionModel
end;
//------------------------------------------------------------------------------ ��������� ����������� �������� �� ��������
procedure TPultUpav.ButtonDinamicClick(Sender: TObject);
begin
DinamicButtonActiv;// FunctionModel
end;
//------------------------------------------------------------------------------ ����� �� ����� ���������� ������ ����������
procedure TPultUpav.ComboBoxDimografiaClick(Sender: TObject);
begin
ComboBoxDimografiaClickk;//Dimografia
end;
//------------------------------------------------------------------------------ ����� �� ����� ���������� ������ �������� �� ��������
procedure TPultUpav.ComboBoxDinamicClick(Sender: TObject);
begin
ComboBoxDinamicClickk;//DinamicPoOtrasl
end;
//------------------------------------------------------------------------------ ����� �� ����� ������� ������ �������� �� ��������
procedure TPultUpav.DBLookupComboBoxDinamicClick(Sender: TObject);
begin
DBLookupComboBoxDinamicClickk;//DinamicPoOtrasl
end;
//------------------------------------------------------------------------------ �������� �����
procedure TPultUpav.FormCreate(Sender: TObject);
begin
SettingCreate;// FunctionModel
end;
procedure TPultUpav.FormDestroy(Sender: TObject);
begin
//FXlsApp.Quit;
end;

//------------------------------------------------------------------------------

end.
