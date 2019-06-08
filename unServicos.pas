unit unServicos;

interface

uses IdHTTP, IdSSLOpenSSL, Classes, System.JSON, SysUtils, XMLDoc, XMLIntf,
  IdSMTP, IdMessage, IdText, IdAttachmentFile,
  IdExplicitTLSClientServerBase;

type
  TCliente = class(TObject)
  private
    fNome: string;
    fIdentidade: string;
    fCPF: string;
    fTelefone: string;
    fEmail: string;
    fCep: string;
    fLogradouro: string;
    fNumero: string;
    fComplemento: string;
    fBairro: string;
    fCidade: string;
    fEstado: string;
    fPais: string;
  public
    property Nome: string read fNome write fNome;
    property Identidade: string read fIdentidade write fIdentidade;
    property CPF: string read fCPF write fCPF;
    property Telefone: string read fTelefone write fTelefone;
    property Email: string read fEmail write fEmail;
    property Cep: string read fCep write fCep;
    property Logradouro: string read fLogradouro write fLogradouro;
    property Numero: string read fNumero write fNumero;
    property Complemento: string read fComplemento write fComplemento;
    property Bairro: string read fBairro write fBairro;
    property Cidade: string read fCidade write fCidade;
    property Estado: string read fEstado write fEstado;
    property Pais: string read fPais write fPais;

end;

type
  TArquivoHTML = class(TObject)
  private
    XMLDocument: TXMLDocument;
    NodeTabela, NodeRegistro, NodeEndereco: IXMLNode;
    fRegistros: integer;
  public
    property TotalRegistros: integer read fRegistros;
    constructor CreateNew(AOwner: TComponent);
    destructor Destroy; override;
    procedure AdicionarCliente(cliente: TCliente);
    function GerarArquivo(): string;

end;

type
  TEmail = class(TObject)
  private
    fDe: string;
    fPara: string;
    fAssunto: String;
    fSenha: string;
    fCorpo: string;
    fArquivo: string;
    IdSSLIOHandlerSocket: TIdSSLIOHandlerSocketOpenSSL;
    IdSMTP: TIdSMTP;
    IdMessage: TIdMessage;
    IdText: TIdText;
    procedure Configuracoes();
  public
    property De: string read fDe write fDe;
    property Para: string read fPara write fPara;
    property Assunto: string read fPara write fPara;
    property Corpo: string read fCorpo write fCorpo;
    property Senha: string read fSenha write fSenha;
    property CaminhoArquivo: string read fArquivo write fArquivo;

    constructor CreateNew(AOwner: TComponent);
    destructor Destroy; override;

    function EnviarEmail(): string;
end;

type
  TServicos = class(TObject)
  private
  public
    class function ConsultaCEP(cep: string): TJSONObject;
    class function ValidarEMail(emails: string): Boolean; static;
end;

implementation

{ TServicos }

class function TServicos.ConsultaCEP(cep: string): TJSONObject;
var
  retorno : TStringStream;
  http: TIdHTTP;
  handler: TIdSSLIOHandlerSocketOpenSSL;
  url: string;
begin
  retorno := TStringStream.Create();
  handler := TIdSSLIOHandlerSocketOpenSSL.Create(nil);
  http := TIdHttp.Create(nil);
  Result:= nil;
  url := 'https://viacep.com.br/ws/' + StringReplace(cep,'-','',[rfReplaceAll]) + '/json/';
  try
    http.ConnectTimeout := 5000;
    http.ReadTimeout := 5000;
    http.MaxAuthRetries := 0;
    http.HTTPOptions := [hoInProcessAuth];
    http.IOHandler := handler;
    http.Request.CustomHeaders.Clear;
    http.Request.BasicAuthentication := true;
    http.Request.UserAgent := 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.96 Safari/537.36';
    try
      http.GET(url, Retorno);
      if (http.ResponseCode = 200) and
         (not(Utf8ToAnsi(retorno.DataString) = '{'#$A'  "erro": true'#$A'}')) then
        Result:= TJSONObject.ParseJSONValue(TEncoding.ASCII.GetBytes( Utf8ToAnsi(retorno.DataString)), 0) as TJSONObject;
    except
      Result:= nil;
    end;
  finally
    FreeAndNil(retorno);
    FreeAndNil(handler);
    FreeAndNil(http);
  end;
end;

class function TServicos.ValidarEMail(emails: string): Boolean;
var
  email: string;
begin
  Result := True;
  emails := Trim(UpperCase(emails));
  if emails = '' then
    Abort();

  if (pos(';',emails) = 0) or (copy(emails,length(emails),1) <> ';') then
    emails := emails + ';';

  email := trim(copy(emails,1,pos(';',emails) - 1));
  emails := copy(emails,pos(';',emails) + 1, length(emails));

  repeat

    if Pos('@', email) > 1 then
    begin
      Delete(email, 1, pos('@', email));
      Result := (Length(email) > 0) and (Pos('.', email) > 2);
    end else
      Result := False;

    if emails = '' then
      email := ''
    else
    begin
      email := trim(copy(emails,1,pos(';',emails) - 1));
      emails := copy(emails,pos(';',emails) + 1, length(emails));
    end;
  until ((email = '') or (Result = False));
end;

{ TCliente }

procedure TArquivoHTML.AdicionarCliente(cliente: TCliente);
begin

  NodeRegistro := NodeTabela.AddChild('Cliente');
  NodeRegistro.ChildValues['Nome'] := cliente.Nome;
  NodeRegistro.ChildValues['Identidade'] := cliente.Identidade;
  NodeRegistro.ChildValues['CPF'] := cliente.CPF;
  NodeRegistro.ChildValues['Telefone'] := cliente.Telefone;
  NodeRegistro.ChildValues['Email'] := cliente.Email;
  NodeEndereco := NodeRegistro.AddChild('Endereco');
  NodeEndereco.ChildValues['Cep'] := cliente.Cep;
  NodeEndereco.ChildValues['Logradouro'] := cliente.Logradouro;
  NodeEndereco.ChildValues['Numero'] := cliente.Numero;
  NodeEndereco.ChildValues['Complemento'] := cliente.Complemento;
  NodeEndereco.ChildValues['Bairro'] := cliente.Bairro;
  NodeEndereco.ChildValues['Cidade'] := cliente.Cidade;
  NodeEndereco.ChildValues['Estado'] := cliente.Estado;
  NodeEndereco.ChildValues['Pais'] := cliente.Pais;
  inc(fRegistros);
end;

constructor TArquivoHTML.CreateNew(AOwner: TComponent);
begin
  inherited;
  XMLDocument := TXMLDocument.Create(nil);
  XMLDocument.Active := True;
  NodeTabela := XMLDocument.AddChild('Clientes');
  fRegistros := 0;
end;

destructor TArquivoHTML.Destroy;
begin
  XMLDocument.Free;
  inherited;
end;

function TArquivoHTML.GerarArquivo: string;
begin
  Result := 'c:\windows\temp\ClientesCadastrados.html';
  XMLDocument.SaveToFile(Result);
end;

{ TEmail }

procedure TEmail.Configuracoes;
begin
    IdSSLIOHandlerSocket.SSLOptions.Method := sslvSSLv23;
    IdSSLIOHandlerSocket.SSLOptions.Mode := sslmClient;

    // Configuração do servidor SMTP (TIdSMTP)
    IdSMTP.IOHandler := IdSSLIOHandlerSocket;
    IdSMTP.UseTLS := utUseImplicitTLS;
    IdSMTP.AuthType := satDefault;
    IdSMTP.Port := 465;
    IdSMTP.Host := 'smtp.gmail.com';
    IdSMTP.Username := fDe;
    IdSMTP.Password := fSenha;

end;

constructor TEmail.CreateNew(AOwner: TComponent);
begin
  inherited;
  IdSSLIOHandlerSocket := TIdSSLIOHandlerSocketOpenSSL.Create(nil);
  IdSMTP := TIdSMTP.Create(nil);
  IdMessage := TIdMessage.Create(nil);
end;

destructor TEmail.Destroy;
begin
  FreeAndNil(IdSSLIOHandlerSocket);
  FreeAndNil(IdSMTP);
  FreeAndNil(IdMessage);
  inherited;
end;

function TEmail.EnviarEmail: string;
var
  erro: string;
  destinatario: string;
begin
  Result := '';
  erro := '';
  if trim(fDe) = '' then
    erro := '- Não foi informado um e-mail de envio' + #13#10;
  if trim(fPara) = '' then
    erro := '- Não foi informado o destinatário do e-mail ' + #13#10;
  if not FileExists(CaminhoArquivo) then
    erro := '- Arquivo não foi localizado';
  if erro = '' then
  begin
    Configuracoes();
    // Configuração da mensagem (TIdMessage)
    IdMessage.From.Address := fDe;
    IdMessage.ReplyTo.EMailAddresses := IdMessage.From.Address;
    destinatario := trim(fPara);
    while pos(';',destinatario) > 0 do
    begin
      IdMessage.Recipients.Add.Text := copy(destinatario,1,pos(';',destinatario) - 1);
      destinatario := trim(copy(destinatario,pos(';',destinatario) + 1, length(destinatario)));
    end;
    if destinatario <> '' then
      IdMessage.Recipients.Add.Text := destinatario;
    IdMessage.Subject := fAssunto;
    IdMessage.Encoding := meMIME;

    // Configuração do corpo do email (TIdText)
    IdText := TIdText.Create(IdMessage.MessageParts);
    IdText.Body.Add(fCorpo);
    IdText.ContentType := 'text/plain; charset=iso-8859-1';

    // Anexo da mensagem (TIdAttachmentFile)
    TIdAttachmentFile.Create(IdMessage.MessageParts, fArquivo);

    try
     // Conexão e autenticação
      try
        IdSMTP.Connect;
        IdSMTP.Authenticate;
      except
        on E:Exception do
        begin
          Result := '- Erro na conexão ou autenticação: ' + E.Message;
          Exit;
        end;
      end;

      // Envio da mensagem
      try
        IdSMTP.Send(IdMessage);
      except
        On E:Exception do
        begin
          Result := '- Erro ao enviar a mensagem: ' + E.Message;
        end;
      end;
    finally
     // desconecta do servidor
     IdSMTP.Disconnect;
     // liberação da DLL
     UnLoadOpenSSLLibrary;
    end;
  end else
    Result := erro;
end;

end.
