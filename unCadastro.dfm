object frmCadCli: TfrmCadCli
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Cadastro de cliente'
  ClientHeight = 428
  ClientWidth = 461
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object tntPaginas: TTabbedNotebook
    Left = 0
    Top = 0
    Width = 461
    Height = 428
    Align = alClient
    TabFont.Charset = DEFAULT_CHARSET
    TabFont.Color = clBtnText
    TabFont.Height = -11
    TabFont.Name = 'Tahoma'
    TabFont.Style = []
    TabOrder = 0
    object TTabPage
      Left = 4
      Top = 24
      Caption = 'Lista de Clientes'
      object Panel1: TPanel
        Left = 0
        Top = 359
        Width = 453
        Height = 41
        Align = alBottom
        TabOrder = 0
        object spbEnviarEmail: TSpeedButton
          Left = 344
          Top = 3
          Width = 91
          Height = 34
          Hint = 'Enviar email dos registros selecionadoss'
          Glyph.Data = {
            B60C0000424DB60C000000000000360000002800000021000000200000000100
            180000000000800C0000C40E0000C40E00000000000000000000FFFFFFFFFFFF
            FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
            FFFFFFFFFFFFFFF6FAF4F1F8EDF8FCF5FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
            FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00FFFF
            FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
            FFFFFFFFFFD0EABFE1F8D2A5E37C8ED95AA1DB79DAEDCDD1EBC0FFFFFFFFFFFF
            FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
            FF00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
            FFFFFFFFFFFFC7F7A895E262A1EE6DAAF678ACF67BA1E9719BE56988D059B8E1
            9DFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
            FFFFFFFFFF00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
            FFFFFFFFFFFFFFCAF8ABA4F36FA9F677A7F675A9F478AAF479ABF57AACF67BAA
            F47989CC5CBBE2A0FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
            FFFFFFFFFFFFFFFFFF00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
            FFFFFFFFFFFFFFFFCDF8B1A1F26BA8F577BCE890CADDA4ADEF7CA8F576A9F577
            ABF47AABF47AADF87C88CE59BEE3A4FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
            FFFFFFFFFFFFFFFFFFFFFFFFFF00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
            FFFFFFFFFFFFFFFFFFF8FEF5C0F69BA8F375A8F475BFF399F1E1E0E8D5CEBAE4
            8EB4E986A6F773A9F478ABF57AA2EC70BBE0A0FFFFFFFFFFFFFFFFFFFFFFFFFF
            FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00FFFFFFFFFFFFFFFFFFFFFFFFFFFF
            FFFFFFFFFFFFFFFFFFFFFFFFFFE3FBD3A1F26AABF47AA8F475C0F199FDEAF5F5
            EDE9EADFD5E0DEC5C3DA98AEEF7EA6F873A8F1779DDF71E5F4DAFFFFFFFFFFFF
            FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00FFFFFFFFFFFFFFFFFFFF
            FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE8FCD7A4F36FABF47AAAF479ACF37C
            AAF479DEEDC7F3ECE6F1EAE2F0E5E1D8D7B7D5D0B1A8F97596E165DEF0CFFFFF
            FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00EAF6FEDCF0FE
            E0F2FEE5F4FEE4F4FEE4F4FEE4F4FEE4F4FEE6F3FFD4F4D4AAF66CABF47AAAF4
            79A9F477D6EDBBF5EAE9F0EAE1F4EAE7EEEBDEEFEDE1A9F679ABF5799DE463CB
            E8CDE7F5FFE4F4FEE4F4FEE4F4FEE4F4FEE5F4FEE0F2FEDDF1FEE7F5FE0054B8
            FC0495FA008CF90797FA1B9FFA189EFA179EFA179EFA149CFD33ADE59BEA89AD
            F577A8F475C0F199F9E9EEF4EAE6F4EAE8D5EEB9BDF196A1F56BAAF478AFF777
            95E1852EA6E1159DFE179EFA179EFA199EFA1BA0FA0797FA008DF90093F936AC
            FB006EC3FD66BFFC1DA0FA0093F91B9FFA2BA5FB2EA7FB2DA7FB2DA7FB219FFF
            74D2B2B5FA6FA8F475C0F199FFE8F9D1EEB3C2F09EA2F56DA8F476ABF47AABF4
            7AB1F46C6AC4AB22A2FF2DA7FB2DA7FB2FA8FB22A2FA1A9FFA0093F91FA1FA51
            B6FB4DB5FC006BC2FD6EC3FD65BFFC67C0FC0092F90E99FA24A2FB2FA8FB2DA7
            FB2BA5FC2CA6FB98E88DB0F773ADF37DB6F28AAAF379A7F474ABF47AABF47AAB
            F47AB2F8729DEA8629A3F92BA6FC2DA7FB2FA8FB24A2FB0C98F90092F967C0FC
            68C0FC5ABBFB4AB4FC006BC2FD6BC2FD6CC2FD6EC3FD6DC2FC0C99F90495F922
            A2FA2AA6FA2EA7FB23A1FF6FCFB6A3EF81B1F871AAF476AAF479ABF47AABF47A
            ADF577B4F9709DEB8825A0FE2AA5FE2EA7FB2AA5FA22A2FA0495F90D99FA6DC2
            FC6EC3FD6EC3FD57B9FB4AB4FC006BC2FD6BC2FD6BC2FD6BC2FD6FC3FD5DBBFC
            30A8FB0C99FA0D99F927A4FA2FA7FB22A0FF39AEEE8CE198AFF774B5FA6FB6FA
            70B7FB6D9AE98A81DAA424A1FF2AA5FE2FA7FB27A4FA0B98F90C99FA30A8FB5E
            BCFC6FC3FD6BC2FD6DC3FD57B9FB4AB4FC006BC2FD6BC2FD6BC2FD6BC2FD6BC2
            FD6DC3FD75C6FD56B9FC0094F90B98F927A4FB31A8FB28A4FF2EA7F959C2CF61
            C5C94ABBC964C7C743B4E5219FFF2BA6FC31A8FB27A3FB0997F90093F956B9FC
            75C6FD6DC3FD6BC2FD6BC2FD6DC3FD57B9FB4AB4FC006BC2FD6BC2FD6BC2FD6B
            C2FD6BC2FD6BC2FD6BC2FD71C4FD5EBCFC2AA6FA0595F90F9AFA31A8FB30A8FC
            1A9DFF0092FF008FFF0795FF1C9EFD32A9FB31A8FB0F9AFA0394F942AFFB61BD
            FC70C4FD6BC2FD6BC2FD6BC2FD6BC2FD6DC3FD57B9FB4AB4FC006BC2FD6BC2FD
            6BC2FD6BC2FD6BC2FD6BC2FD6BC2FD6BC2FD6DC3FD77C7FD47B2FB0A97F90796
            F90495F90A97F978C7FD9BD6FF1C9FFA109AFA0395F90796F90A97F947B2FB75
            C6FD6DC2FD6BC2FD6BC2FD6BC2FD6BC2FD6BC2FD6DC3FD57B9FB4AB4FC006BC2
            FD6BC2FD6BC2FD6BC2FD6BC2FD6BC2FD6BC2FD6BC2FD6BC2FD6CC2FD6DC3FC30
            A8FB20A1FA3DADFB6EC3FC92D2FE8BCFFE9AD5FE6BC1FC3BACFB20A1FA30A8FB
            6FC4FC6EC3FD6BC2FD6BC2FD6BC2FD6BC2FD6BC2FD6BC2FD6DC3FD57B9FB4AB4
            FC006BC2FD6BC2FD6BC2FD6BC2FD6BC2FD6BC2FD6BC2FD6DC3FD6CC2FD52B8FC
            44B1FB6AC1FC79C8FD95D4FE8ED0FE8ACFFE8ACFFE8ACFFE8ED0FE95D4FE79C8
            FD6AC1FC46B2FB54B8FC6CC2FD6DC3FD6BC2FD6BC2FD6BC2FD6BC2FD6DC3FD57
            B9FB4AB4FC006BC2FD6BC2FD6BC2FD6BC2FD6BC2FD6CC2FD71C4FD55B8FB48B3
            FB55B9FC84CCFD8ACFFD8DD0FE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8A
            CFFE8DD0FE8BCFFD7ECAFD51B8FC4AB4FC55B8FB71C4FD6CC2FD6BC2FD6BC2FD
            6DC3FD57B9FB4AB4FC006BC2FD6BC2FD6BC2FD6CC2FD67C0FC5BBBFC41B0FB4E
            B6FC67C0FD93D3FE8BCFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE
            8ACFFE8ACFFE8ACFFE8ACFFE8CCFFE95D3FE58BAFC4FB6FC41B0FB5BBBFC69C1
            FC6CC2FD6DC3FD57B9FB4AB4FC006BC2FD6BC2FD6EC2FD63BFFC42B1FB4EB6FC
            61BEFC87CEFE8FD1FE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACF
            FE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE90D1FE87CEFE61BEFC4E
            B6FC4AB4FB63BEFC70C4FD58B9FB4AB4FC006EC3FD65BFFD53B7FC40B0FB74C5
            FD7CC9FD90D1FE8BCFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8A
            CFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8BCFFE
            90D1FE7CC9FD73C5FD41B0FB57B9FC50B6FB4BB4FC004EB6FC46B2FB66C0FC83
            CCFD8ED0FE8CCFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE
            8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACF
            FE8ACFFE8ACFFE8CCFFE8ED0FE7DCAFD55B9FC45B2FB4AB4FB0050B6FC69C1FD
            92D2FE8BCFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACF
            FE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8A
            CFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8CCFFE93D2FE6AC1FD51B7FC0090D1
            FE8DD0FE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8A
            CFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE
            8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8DD0FE90D1
            FE008ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE
            8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACF
            FE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8A
            CFFE8ACFFE008ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACF
            FE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8A
            CFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE
            8ACFFE8ACFFE8ACFFE0088CEFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8A
            CFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE
            8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACFFE8ACF
            FE8ACFFE8ACFFE8ACFFE85CDFD0092D2FD7FCAFD81CBFD81CBFD81CBFD81CBFD
            81CBFD81CBFD81CBFD81CBFD81CBFD81CBFD81CBFD81CBFD81CBFD81CBFD81CB
            FD81CBFD81CBFD81CBFD81CBFD81CBFD81CBFD81CBFD81CBFD81CBFD81CBFD81
            CBFD81CBFD81CBFD81CBFD7CC9FDB0DEFE00E4F3FED1ECFED3ECFED3ECFED3EC
            FED3ECFED3ECFED3ECFED3ECFED3ECFED3ECFED3ECFED3ECFED3ECFED3ECFED3
            ECFED3ECFED3ECFED3ECFED3ECFED3ECFED3ECFED3ECFED3ECFED3ECFED3ECFE
            D3ECFED3ECFED3ECFED3ECFED3ECFECEEBFEFFFFFF00FFFFFFFFFFFFFFFFFFFF
            FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
            FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
            FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00}
          OnClick = spbEnviarEmailClick
        end
        object DBNavigator2: TDBNavigator
          Left = 7
          Top = 3
          Width = 330
          Height = 34
          DataSource = dtsCliente
          VisibleButtons = [nbInsert, nbDelete, nbEdit]
          Hints.Strings = (
            'Primeiro registro'
            'Registro anterior'
            'Pr'#243'ximo registro'
            #218'ltimo registro'
            'Novo cadastro'
            'Remover cadastro'
            'Alterar cadastro'
            'Salvar'
            'Cancelar altera'#231#227'o'
            'Atualizar dados')
          TabOrder = 0
          OnClick = DBNavigator2Click
        end
      end
      object grdClientes: TDBGrid
        Left = 0
        Top = 0
        Width = 453
        Height = 359
        Align = alClient
        DataSource = dtsCliente
        Options = [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgMultiSelect, dgTitleClick, dgTitleHotTrack]
        ReadOnly = True
        TabOrder = 1
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        Columns = <
          item
            Expanded = False
            FieldName = 'CPF'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Nome'
            Width = 220
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Telefone'
            Visible = True
          end>
      end
    end
    object TTabPage
      Left = 4
      Top = 24
      Caption = 'Cadastro'
      object Label1: TLabel
        Left = 51
        Top = 10
        Width = 27
        Height = 13
        Caption = 'Nome'
      end
      object Identidade: TLabel
        Left = 28
        Top = 34
        Width = 52
        Height = 13
        Caption = 'Identidade'
      end
      object CPF: TLabel
        Left = 59
        Top = 58
        Width = 19
        Height = 13
        Caption = 'CPF'
      end
      object Telefone: TLabel
        Left = 36
        Top = 82
        Width = 42
        Height = 13
        Caption = 'Telefone'
      end
      object Email: TLabel
        Left = 54
        Top = 106
        Width = 24
        Height = 13
        Caption = 'Email'
      end
      object Endereço: TGroupBox
        Left = 8
        Top = 125
        Width = 425
        Height = 214
        Caption = 'Endere'#231'o'
        TabOrder = 5
        object Label6: TLabel
          Left = 51
          Top = 18
          Width = 19
          Height = 13
          Caption = 'Cep'
        end
        object Label7: TLabel
          Left = 15
          Top = 42
          Width = 55
          Height = 13
          Caption = 'Logradouro'
        end
        object Label8: TLabel
          Left = 33
          Top = 66
          Width = 37
          Height = 13
          Caption = 'N'#250'mero'
        end
        object Label9: TLabel
          Left = 5
          Top = 90
          Width = 65
          Height = 13
          Caption = 'Complemento'
        end
        object Label10: TLabel
          Left = 42
          Top = 114
          Width = 28
          Height = 13
          Caption = 'Bairro'
        end
        object Label11: TLabel
          Left = 37
          Top = 138
          Width = 33
          Height = 13
          Caption = 'Cidade'
        end
        object Label12: TLabel
          Left = 37
          Top = 162
          Width = 33
          Height = 13
          Caption = 'Estado'
        end
        object Label13: TLabel
          Left = 51
          Top = 186
          Width = 19
          Height = 13
          Caption = 'Pa'#237's'
        end
        object lblCepNaoLocalizado: TLabel
          Left = 160
          Top = 18
          Width = 104
          Height = 13
          Caption = 'Cep n'#227'o localizado'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clMaroon
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          ParentFont = False
          Visible = False
        end
        object edtLogradouro: TDBEdit
          Left = 76
          Top = 40
          Width = 343
          Height = 21
          DataField = 'Logradouro'
          DataSource = dtsCliente
          TabOrder = 1
        end
        object edtNumero: TDBEdit
          Left = 76
          Top = 64
          Width = 77
          Height = 21
          DataField = 'Numero'
          DataSource = dtsCliente
          TabOrder = 2
        end
        object edtComplemento: TDBEdit
          Left = 76
          Top = 87
          Width = 343
          Height = 21
          DataField = 'Complemento'
          DataSource = dtsCliente
          TabOrder = 3
        end
        object edtBairro: TDBEdit
          Left = 76
          Top = 112
          Width = 221
          Height = 21
          DataField = 'Bairro'
          DataSource = dtsCliente
          TabOrder = 4
        end
        object edtCidade: TDBEdit
          Left = 76
          Top = 136
          Width = 221
          Height = 21
          DataField = 'Cidade'
          DataSource = dtsCliente
          TabOrder = 5
        end
        object edtEstado: TDBEdit
          Left = 76
          Top = 160
          Width = 221
          Height = 21
          DataField = 'Estado'
          DataSource = dtsCliente
          TabOrder = 6
        end
        object edtPais: TDBEdit
          Left = 76
          Top = 184
          Width = 221
          Height = 21
          DataField = 'Pais'
          DataSource = dtsCliente
          TabOrder = 7
        end
        object edtCep: TDBEdit
          Left = 76
          Top = 15
          Width = 77
          Height = 21
          DataField = 'CEP'
          DataSource = dtsCliente
          TabOrder = 0
          OnChange = edtCepChange
        end
        object TabSet1: TTabSet
          Left = 224
          Top = -72
          Width = 185
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
        end
      end
      object edtNome: TDBEdit
        Left = 84
        Top = 8
        Width = 343
        Height = 21
        DataField = 'Nome'
        DataSource = dtsCliente
        TabOrder = 0
      end
      object edtIdentidade: TDBEdit
        Left = 84
        Top = 32
        Width = 93
        Height = 21
        DataField = 'Identidade'
        DataSource = dtsCliente
        TabOrder = 1
      end
      object edtEmail: TDBEdit
        Left = 84
        Top = 104
        Width = 343
        Height = 21
        DataField = 'Email'
        DataSource = dtsCliente
        TabOrder = 4
        OnExit = edtEmailExit
      end
      object edtCpf: TDBEdit
        Left = 84
        Top = 56
        Width = 93
        Height = 21
        DataField = 'CPF'
        DataSource = dtsCliente
        TabOrder = 2
      end
      object edtTelefone: TDBEdit
        Left = 84
        Top = 79
        Width = 93
        Height = 21
        DataField = 'Telefone'
        DataSource = dtsCliente
        TabOrder = 3
      end
      object Panel3: TPanel
        Left = 0
        Top = 359
        Width = 453
        Height = 41
        Align = alBottom
        TabOrder = 6
        object DBNavigator1: TDBNavigator
          Left = 3
          Top = 2
          Width = 432
          Height = 34
          DataSource = dtsCliente
          VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast, nbInsert, nbDelete, nbEdit, nbPost, nbCancel]
          Hints.Strings = (
            'Primeiro registro'
            'Registro anterior'
            'Pr'#243'ximo registro'
            #218'ltimo registro'
            'Novo cadastro'
            'Remover cadastro'
            'Alterar cadastro'
            'Salvar'
            'Cancelar altera'#231#227'o'
            'Atualizar dados')
          TabOrder = 0
          OnClick = DBNavigator1Click
        end
      end
    end
  end
  object tblCliente: TFDMemTable
    BeforePost = tblClienteBeforePost
    FetchOptions.AssignedValues = [evMode]
    FetchOptions.Mode = fmAll
    ResourceOptions.AssignedValues = [rvSilentMode]
    ResourceOptions.SilentMode = True
    UpdateOptions.AssignedValues = [uvCheckRequired, uvAutoCommitUpdates]
    UpdateOptions.CheckRequired = False
    UpdateOptions.AutoCommitUpdates = True
    Left = 264
    Top = 64
    object tblClienteNome: TStringField
      FieldName = 'Nome'
      Required = True
      Size = 120
    end
    object tblClienteIdentidade: TStringField
      FieldName = 'Identidade'
      Size = 15
    end
    object tblClienteCPF: TStringField
      FieldName = 'CPF'
      EditMask = '000\.000\.000\-00;1;_'
      Size = 14
    end
    object tblClienteTelefone: TStringField
      FieldName = 'Telefone'
      EditMask = '!\(99\)90000-0000;1;_'
      Size = 15
    end
    object tblClienteEmail: TStringField
      FieldName = 'Email'
      Size = 200
    end
    object tblClienteCEP: TStringField
      FieldName = 'CEP'
      EditMask = '00000\-999;1;_'
      Size = 9
    end
    object tblClienteLogradouro: TStringField
      DisplayWidth = 120
      FieldName = 'Logradouro'
      Size = 120
    end
    object tblClienteNumero: TStringField
      FieldName = 'Numero'
      Size = 15
    end
    object tblClienteComplemento: TStringField
      FieldName = 'Complemento'
      Size = 120
    end
    object tblClienteBairro: TStringField
      FieldName = 'Bairro'
      Size = 50
    end
    object tblClienteCidade: TStringField
      FieldName = 'Cidade'
      Size = 50
    end
    object tblClienteEstado: TStringField
      FieldName = 'Estado'
      Size = 50
    end
    object tblClientePais: TStringField
      FieldName = 'Pais'
      Size = 50
    end
  end
  object dtsCliente: TDataSource
    DataSet = tblCliente
    Left = 303
    Top = 64
  end
end
