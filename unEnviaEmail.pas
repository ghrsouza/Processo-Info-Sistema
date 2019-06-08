unit unEnviaEmail;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Mask, Vcl.DBCtrls, unServicos,
  Vcl.Buttons, unCadastro;

type
  TfrmEnviaEmail = class(TForm)
    Label1: TLabel;
    Identidade: TLabel;
    CPF: TLabel;
    edtDe: TEdit;
    edtPara: TEdit;
    edtAssunto: TEdit;
    GroupBox1: TGroupBox;
    lblAnexo: TLabel;
    memCorpo: TMemo;
    btnEnviar: TSpeedButton;
    Label2: TLabel;
    edtSenha: TEdit;
    procedure edtDeExit(Sender: TObject);
    procedure edtParaExit(Sender: TObject);
    procedure btnEnviarClick(Sender: TObject);
  private
    caminhoArquivo: string;
    { Private declarations }
  public

    constructor Create(AOwner: TComponent; arquivo: String; qtdRegistros: Integer); overload;
    procedure ObterDados(arquivo: String; qtdRegistros: Integer);

    { Public declarations }
  end;

var
  frmEnviaEmail: TfrmEnviaEmail;

implementation

{$R *.dfm}

constructor TfrmEnviaEmail.Create(AOwner: TComponent; arquivo: String;
  qtdRegistros: Integer);
begin
  self.Create(AOwner);
  caminhoArquivo := arquivo;
  if qtdRegistros = 1 then
    lblAnexo.Caption := 'Segue anexo arquivo com 1 registro de cliente'
  else
    lblAnexo.Caption := StringReplace('Segue anexo arquivo com XX registro de cliente','XX',inttostr(qtdRegistros), [rfReplaceAll]);
end;

procedure TfrmEnviaEmail.edtDeExit(Sender: TObject);
begin
  if (not TServicos.ValidarEMail(EdtDe.Text)) or (not Pos('@GMAIL.', EdtDe.Text) = 0) then
  begin
    MessageDlg('E-mail de envio não é válido. Favor informar um email do GMAIL.', mtinformation,
      [mbok],0);
    edtDe.SetFocus();
  end;

end;

procedure TfrmEnviaEmail.edtParaExit(Sender: TObject);
begin
  if not TServicos.ValidarEMail(edtPara.Text) then
  begin
    MessageDlg('Encontrado e-mail incorreto, favor corrigir.', mtinformation,
      [mbok],0);
    edtPara.SetFocus();
  end;

end;

procedure TfrmEnviaEmail.ObterDados(arquivo: String; qtdRegistros: Integer);
begin
  caminhoArquivo := arquivo;
  if qtdRegistros = 1 then
    lblAnexo.Caption := 'Segue anexo arquivo com 1 registro de cliente'
  else
    lblAnexo.Caption := StringReplace('Segue anexo arquivo com XX registro de cliente','XX',inttostr(qtdRegistros), [rfReplaceAll]);
end;

procedure TfrmEnviaEmail.btnEnviarClick(Sender: TObject);
var
  email: TEmail;
  retorno: string;
begin
  email := TEmail.CreateNew(nil);
  try
    Screen.Cursor:= crHourGlass;
    btnEnviar.Enabled:= False;

    email.De := edtDe.Text;
    email.Para := edtPara.Text;
    email.Assunto := edtAssunto.Text;
    email.Corpo := memCorpo.Text;
    email.Senha := edtSenha.Text;
    email.CaminhoArquivo := caminhoArquivo;
    retorno := email.EnviarEmail();
    if retorno <> '' then
    begin
      MessageDlg('O e-mail não foi possível ser enviado. ' + #13#10 +
                retorno,mtinformation,[mbok],0);
    end else
      MessageDlg('Mensagem enviada com sucesso!',mtInformation,[mbOK],0);
  finally
    FreeAndNil(email);
    btnEnviar.Enabled:= True;
    Screen.Cursor:= crDefault;
  end;

end;

end.
