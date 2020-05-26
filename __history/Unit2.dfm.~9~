object Dm1: TDm1
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Height = 505
  Width = 794
  object ADOConnection1: TADOConnection
    ConnectionString = 
      'Provider=SQLOLEDB.1;Password=A2014a;Persist Security Info=True;U' +
      'ser ID=sa;Initial Catalog=Data;Data Source=DESKTOP-TRF95A6'
    LoginPrompt = False
    Provider = 'SQLOLEDB.1'
    Left = 256
    Top = 8
  end
  object ADOQuery1: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select ['#1053#1072#1079#1074#1072#1085#1080#1077'] from [dbo].['#1054#1090#1088#1072#1089#1083#1080']')
    Left = 24
    Top = 72
  end
  object DataSource1: TDataSource
    DataSet = ADOQuery1
    Left = 24
    Top = 128
  end
end
