object dm_Image: Tdm_Image
  OldCreateOrder = True
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Height = 392
  Width = 374
  object PrintDialog: TPrintDialog
    Collate = True
    FromPage = 1
    MaxPage = 99999
    Options = [poPageNums, poWarning, poHelp]
    ToPage = 1
    Left = 176
    Top = 112
  end
end
