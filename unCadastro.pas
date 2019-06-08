unit unCadastro;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Mask, Vcl.ExtCtrls, Vcl.StdCtrls,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdHTTP,
  dbxjson, Vcl.ComCtrls, Vcl.Tabs, Vcl.TabNotBk, Vcl.DBCtrls, FireDAC.Stan.Intf,
  FireDAC.Stan.Option, FireDAC.Stan.Param, FireDAC.Stan.Error, FireDAC.DatS,
  FireDAC.Phys.Intf, FireDAC.DApt.Intf, Data.DB, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client, Vcl.Grids, Vcl.DBGrids, Vcl.Buttons;

type
  TfrmCadCli = class(TForm)
    tntPaginas: TTabbedNotebook;
    Panel1: TPanel;
    Label1: TLabel;
    Identidade: TLabel;
    CPF: TLabel;
    Telefone: TLabel;
    Email: TLabel;
    Endereço: TGroupBox;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    edtLogradouro: TDBEdit;
    edtNumero: TDBEdit;
    edtComplemento: TDBEdit;
    edtBairro: TDBEdit;
    edtCidade: TDBEdit;
    edtEstado: TDBEdit;
    edtPais: TDBEdit;
    edtCep: TDBEdit;
    TabSet1: TTabSet;
    edtNome: TDBEdit;
    edtIdentidade: TDBEdit;
    edtEmail: TDBEdit;
    edtCpf: TDBEdit;
    edtTelefone: TDBEdit;
    Panel3: TPanel;
    DBNavigator1: TDBNavigator;
    tblCliente: TFDMemTable;
    tblClienteNome: TStringField;
    tblClienteIdentidade: TStringField;
    tblClienteCPF: TStringField;
    tblClienteTelefone: TStringField;
    tblClienteEmail: TStringField;
    tblClienteCEP: TStringField;
    tblClienteLogradouro: TStringField;
    tblClienteNumero: TStringField;
    tblClienteComplemento: TStringField;
    tblClienteBairro: TStringField;
    tblClienteCidade: TStringField;
    tblClienteEstado: TStringField;
    tblClientePais: TStringField;
    dtsCliente: TDataSource;
    DBNavigator2: TDBNavigator;
    grdClientes: TDBGrid;
    lblCepNaoLocalizado: TLabel;
    spbEnviarEmail: TSpeedButton;
    procedure edtCepChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DBNavigator2Click(Sender: TObject; Button: TNavigateBtn);
    procedure edtEmailExit(Sender: TObject);
    procedure spbEnviarEmailClick(Sender: TObject);
    procedure DBNavigator1Click(Sender: TObject; Button: TNavigateBtn);
    procedure tblClienteBeforePost(DataSet: TDataSet);
  private
    { Private declarations }
    procedure CarregarEndereco(cep: string);
  public
    { Public declarations }
  end;

var
  frmCadCli: TfrmCadCli;

implementation

uses
  unServicos, System.JSON, unEnviaEmail;

{$R *.dfm}

procedure TfrmCadCli.CarregarEndereco(cep: string);
var
  retorno: TJSONObject;
begin
  retorno := TServicos.ConsultaCEP(cep);
  lblCepNaoLocalizado.Visible := False;
  if retorno <> nil then
  begin
    tblClienteLogradouro.AsString := retorno.GetValue('logradouro').Value;
    tblClienteComplemento.AsString := retorno.GetValue('complemento').Value;
    tblClienteBairro.AsString := retorno.GetValue('bairro').Value;
    tblClienteCidade.AsString := retorno.GetValue('localidade').Value;
    tblClienteEstado.AsString := retorno.GetValue('uf').Value;
  end else
  begin
    lblCepNaoLocalizado.Visible := True;
  end;
end;

procedure TfrmCadCli.DBNavigator1Click(Sender: TObject; Button: TNavigateBtn);
begin
  if Button in [nbInsert,nbEdit] then
  begin
    tntPaginas.PageIndex := 1;
    edtNome.SetFocus();
  end;

end;

procedure TfrmCadCli.DBNavigator2Click(Sender: TObject; Button: TNavigateBtn);
begin
  if Button <> nbDelete then
  begin
    tntPaginas.PageIndex := 1;
    edtNome.SetFocus();
  end;
end;

procedure TfrmCadCli.edtCepChange(Sender: TObject);
begin
  if (length(trim(edtCep.Text)) = 9) and (tblCliente.State in [dsInsert,dsEdit]) then
    CarregarEndereco(edtCep.Text);

end;

procedure TfrmCadCli.edtEmailExit(Sender: TObject);
begin
  if not TServicos.ValidarEMail(EdtEmail.Text) then
  begin
    MessageDlg('Encontrado e-mail incorreto, favor corrigir.', mtinformation, [mbok],0);
    EdtEmail.SetFocus();
  end;
end;

procedure TfrmCadCli.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  tblCliente.Close();
  Action := caFree;
end;

procedure TfrmCadCli.FormCreate(Sender: TObject);
begin
  tblCliente.Open();
  tntPaginas.PageIndex := 0;

  lblCepNaoLocalizado.Visible := False;

end;

procedure TfrmCadCli.spbEnviarEmailClick(Sender: TObject);
var
  sel: Integer;
  cliente: TCliente;
  arquivo: TArquivoHTML;
  frmEnviaEmail : TfrmEnviaEmail;
begin
  if grdClientes.SelectedRows.Count = 0 then
  begin
    MessageDlg('Favor selecionar pelo menos um cliente.',mtinformation,[mbok],0);
    Abort();
  end;

  arquivo := TArquivoHTML.CreateNew(nil);
  try
    for sel:= 0 to grdClientes.SelectedRows.Count-1 do begin
      //enviar os emails
      tblClienteNome.Dataset.GotoBookmark (grdClientes.SelectedRows.Items [sel]);
      cliente := TCliente.Create();
      try

        cliente.Nome := tblClienteNome.AsString;
        cliente.Identidade := tblClienteIdentidade.AsString;
        cliente.CPF := tblClienteCPF.AsString;
        cliente.Telefone := tblClienteTelefone.AsString;
        cliente.Email := tblClienteEmail.AsString;
        cliente.Cep := tblClienteCEP.AsString;
        cliente.Logradouro := tblClienteLogradouro.AsString;
        cliente.Numero := tblClienteNumero.AsString;
        cliente.Complemento := tblClienteComplemento.AsString;
        cliente.Bairro := tblClienteBairro.AsString;
        cliente.Cidade := tblClienteCidade.AsString;
        cliente.Estado := tblClienteEstado.AsString;
        cliente.Pais := tblClientePais.AsString;

        arquivo.AdicionarCliente(cliente);
      finally
        FreeAndNil(cliente);
      end;
    end;
    frmEnviaEmail := TfrmEnviaEmail.Create(self, arquivo.GerarArquivo(), arquivo.TotalRegistros);
    try
      frmEnviaEmail.ShowModal;
    finally
      FreeAndNil(frmEnviaEmail);
    end;
  finally
    FreeAndNil(arquivo);
  end;
end;

procedure TfrmCadCli.tblClienteBeforePost(DataSet: TDataSet);
begin
  if trim(tblClienteNome.AsString) = '' then
  begin
    MessageDlg('O nome do cliente é obrigatório.',mtInformation,[mbok],0);
    edtNome.SetFocus();
    Abort();
  end;
end;

end.
