#define MyAppName "Windsor Widget"
#define MyAppVersion "1.1.5"
#define MyAppPublisher "Brad Mayze"
#define MyAppExeName "WindsorWidget.exe"

[Setup]
AppId={{E9179E8A-22E8-4B27-93FA-CB12C1F5F3E7}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\Windsor Widget
DefaultGroupName=Windsor Widget
DisableProgramGroupPage=yes
OutputBaseFilename=WindsorWidget_Client_1_1_5
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
SetupIconFile=..\assets\windsor_icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; Flags: unchecked

[Files]
Source: "..\dist\WindsorWidget\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\scripts\test_sql_connection.ps1"; Flags: dontcopy
Source: "..\prereqs\msodbcsql18.msi"; DestDir: "{tmp}"; Flags: deleteafterinstall ignoreversion skipifsourcedoesntexist
Source: "..\prereqs\msodbcsql18.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall ignoreversion skipifsourcedoesntexist

[Icons]
Name: "{autoprograms}\Windsor Widget"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\Windsor Widget"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "msiexec.exe"; Parameters: "/i ""{tmp}\msodbcsql18.msi"" /passive /norestart IACCEPTMSODBCSQLLICENSETERMS=YES"; StatusMsg: "Installing ODBC Driver 18 for SQL Server..."; Flags: waituntilterminated; Check: (not IsOdbc18Installed) and HasBundledOdbcMsi
Filename: "{tmp}\msodbcsql18.exe"; Parameters: "/quiet /norestart IACCEPTMSODBCSQLLICENSETERMS=YES"; StatusMsg: "Installing ODBC Driver 18 for SQL Server..."; Flags: waituntilterminated; Check: (not IsOdbc18Installed) and (not HasBundledOdbcMsi) and HasBundledOdbcExe
Filename: "{app}\{#MyAppExeName}"; Description: "Launch Windsor Widget"; Flags: nowait postinstall skipifsilent; Check: IsOdbc18Installed or HasBundledOdbcMsi or HasBundledOdbcExe

[Code]
var
  SqlPage: TInputQueryWizardPage;
  FolderPage: TInputDirWizardPage;
  YuTemplatePage: TInputDirWizardPage;
  TestButton: TNewButton;
  ConnectionValidated: Boolean;
  LastValidatedSignature: string;

function IsOdbc18Installed: Boolean;
begin
  Result :=
    RegKeyExists(HKLM64, 'SOFTWARE\ODBC\ODBCINST.INI\ODBC Driver 18 for SQL Server') or
    RegKeyExists(HKLM32, 'SOFTWARE\ODBC\ODBCINST.INI\ODBC Driver 18 for SQL Server');
end;

function HasBundledOdbcMsi: Boolean;
begin
  Result := FileExists(ExpandConstant('{tmp}\msodbcsql18.msi'));
end;

function HasBundledOdbcExe: Boolean;
begin
  Result := FileExists(ExpandConstant('{tmp}\msodbcsql18.exe'));
end;

function InitializeSetup(): Boolean;
begin
  Result := True;
  if not IsOdbc18Installed then
  begin
    MsgBox(
      'ODBC Driver 18 for SQL Server is not currently installed.' + #13#10#13#10 +
      'That is fine if you bundled the Microsoft ODBC installer into this build. ' +
      'The installer will attempt to install it before first launch.' + #13#10#13#10 +
      'Python app dependencies are already bundled into the Windsor Widget executable.',
      mbInformation,
      MB_OK);
  end;
end;

function PosFrom(const Needle, Haystack: string; const Offset: Integer): Integer;
var
  TailText: string;
  RelativePos: Integer;
begin
  if Offset <= 1 then
  begin
    Result := Pos(Needle, Haystack);
    exit;
  end;

  TailText := Copy(Haystack, Offset, MaxInt);
  RelativePos := Pos(Needle, TailText);
  if RelativePos > 0 then
    Result := RelativePos + Offset - 1
  else
    Result := 0;
end;

function JsonEscape(const Value: string): string;
var
  Escaped: string;
begin
  Escaped := Value;
  StringChangeEx(Escaped, '\', '\\', True);
  StringChangeEx(Escaped, '"', '\"', True);
  Result := Escaped;
end;

function GetConfigSignature(): string;
begin
  Result :=
    Trim(SqlPage.Values[0]) + '|' +
    Trim(SqlPage.Values[1]) + '|' +
    Trim(SqlPage.Values[2]) + '|' +
    Trim(SqlPage.Values[3]) + '|' +
    Trim(SqlPage.Values[4]);
end;

procedure MarkConnectionDirty(Sender: TObject);
begin
  ConnectionValidated := False;
end;

function ReadExistingConfigValue(const Key: string; const Default: string): string;
var
  JsonPath: string;
  JsonText: AnsiString;
  SearchText: AnsiString;
  StartPos: Integer;
  EndPos: Integer;
begin
  Result := Default;
  JsonPath := ExpandConstant('{commonappdata}\WindsorWidget\client_config.json');
  if not FileExists(JsonPath) then
    exit;

  if not LoadStringFromFile(JsonPath, JsonText) then
    exit;

  SearchText := '"' + Key + '"';
  StartPos := Pos(SearchText, JsonText);
  if StartPos <= 0 then
    exit;

  StartPos := PosFrom(':', JsonText, StartPos);
  if StartPos <= 0 then
    exit;

  StartPos := PosFrom('"', JsonText, StartPos);
  if StartPos <= 0 then
    exit;
  StartPos := StartPos + 1;

  EndPos := PosFrom('"', JsonText, StartPos);
  if EndPos <= StartPos then
    exit;

  Result := Copy(string(JsonText), StartPos, EndPos - StartPos);
end;

procedure TestConnectionButtonClick(Sender: TObject);
var
  Server, Port, Database, Username, Password: string;
  ScriptPath, OutputPath, Params: string;
  MessageText: AnsiString;
  ResultCode: Integer;
begin
  Server := Trim(SqlPage.Values[0]);
  Port := Trim(SqlPage.Values[1]);
  Database := Trim(SqlPage.Values[2]);
  Username := Trim(SqlPage.Values[3]);
  Password := Trim(SqlPage.Values[4]);

  if (Server = '') or (Database = '') or (Username = '') or (Password = '') then
  begin
    MsgBox('Enter server, database, username, and password first.', mbError, MB_OK);
    exit;
  end;

  ExtractTemporaryFile('test_sql_connection.ps1');
  ScriptPath := ExpandConstant('{tmp}\test_sql_connection.ps1');
  OutputPath := ExpandConstant('{tmp}\windsor_sql_test.txt');

  if FileExists(OutputPath) then
    DeleteFile(OutputPath);

  Params :=
    '-ExecutionPolicy Bypass -NoProfile -File ' + AddQuotes(ScriptPath) +
    ' -Server ' + AddQuotes(Server) +
    ' -Port ' + AddQuotes(Port) +
    ' -Database ' + AddQuotes(Database) +
    ' -Username ' + AddQuotes(Username) +
    ' -Password ' + AddQuotes(Password) +
    ' -OutputFile ' + AddQuotes(OutputPath);

  if Exec(ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe'), Params, '', SW_HIDE, ewWaitUntilTerminated, ResultCode) and (ResultCode = 0) then
  begin
    ConnectionValidated := True;
    LastValidatedSignature := GetConfigSignature();
    MsgBox('Connection successful.', mbInformation, MB_OK);
  end
  else
  begin
    ConnectionValidated := False;
    MessageText := 'Connection test failed.';
    if FileExists(OutputPath) and LoadStringFromFile(OutputPath, MessageText) then
      MsgBox(string(MessageText), mbError, MB_OK)
    else
      MsgBox('Connection test failed. Check the server details and try again.', mbError, MB_OK);
  end;
end;

procedure InitializeWizard();
var
  ExistingFolder: string;
  ExistingYuTemplatePath: string;
  ExistingYuTemplateFolder: string;
begin
  SqlPage := CreateInputQueryPage(
    wpSelectDir,
    'Server connection',
    'Enter the SQL Server connection details',
    'Use the Test Connection button before continuing.'
  );
  SqlPage.Add('Server:', False);
  SqlPage.Add('Port:', False);
  SqlPage.Add('Database:', False);
  SqlPage.Add('Username:', False);
  SqlPage.Add('Password:', True);

  SqlPage.Values[0] := ReadExistingConfigValue('server', '');
  SqlPage.Values[1] := ReadExistingConfigValue('port', '14330');
  SqlPage.Values[2] := ReadExistingConfigValue('database', 'WindsorWidget');
  SqlPage.Values[3] := ReadExistingConfigValue('username', 'windsor_app');
  SqlPage.Values[4] := ReadExistingConfigValue('password', '');

  TestButton := TNewButton.Create(SqlPage);
  TestButton.Parent := SqlPage.Surface;
  TestButton.Left := SqlPage.Edits[4].Left;
  TestButton.Top := SqlPage.Edits[4].Top + SqlPage.Edits[4].Height + ScaleY(12);
  TestButton.Width := ScaleX(140);
  TestButton.Height := ScaleY(24);
  TestButton.Caption := 'Test Connection';
  TestButton.OnClick := @TestConnectionButtonClick;

  SqlPage.Edits[0].OnChange := @MarkConnectionDirty;
  SqlPage.Edits[1].OnChange := @MarkConnectionDirty;
  SqlPage.Edits[2].OnChange := @MarkConnectionDirty;
  SqlPage.Edits[3].OnChange := @MarkConnectionDirty;
  SqlPage.Edits[4].OnChange := @MarkConnectionDirty;

  FolderPage := CreateInputDirPage(
    SqlPage.ID,
    'Customer files folder',
    'Choose the local folder that holds the customer files',
    'Only the filename is stored in SQL. This folder is saved locally on this PC.',
    False,
    ''
  );
  FolderPage.Add('Customer files folder:');

  ExistingFolder := ExpandConstant('{reg:HKCU\Software\Windsor\WidgetApp,customerFilesRoot|}');
  if ExistingFolder <> '' then
    FolderPage.Values[0] := ExistingFolder;

  YuTemplatePage := CreateInputDirPage(
    FolderPage.ID,
    'YU template folder',
    'Choose the local folder that contains the YU workbook template',
    'The installer will save the full path to yuchang_order_form_Widget.xlsx for this PC/user.',
    False,
    ''
  );
  YuTemplatePage.Add('YU template folder:');

  ExistingYuTemplatePath := ExpandConstant('{reg:HKCU\Software\Windsor\WidgetApp\yu,template_path|}');
  if ExistingYuTemplatePath <> '' then
  begin
    ExistingYuTemplateFolder := ExtractFileDir(ExistingYuTemplatePath);
    if ExistingYuTemplateFolder <> '' then
      YuTemplatePage.Values[0] := ExistingYuTemplateFolder;
  end;

  ConnectionValidated := False;
  LastValidatedSignature := '';
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;

  if CurPageID = SqlPage.ID then
  begin
    if (Trim(SqlPage.Values[0]) = '') or (Trim(SqlPage.Values[2]) = '') or
       (Trim(SqlPage.Values[3]) = '') or (Trim(SqlPage.Values[4]) = '') then
    begin
      MsgBox('Server, database, username, and password are required.', mbError, MB_OK);
      Result := False;
      exit;
    end;

    if (not ConnectionValidated) or (LastValidatedSignature <> GetConfigSignature()) then
    begin
      TestConnectionButtonClick(nil);
      if not ConnectionValidated then
      begin
        Result := False;
        exit;
      end;
    end;
  end;

  if CurPageID = FolderPage.ID then
  begin
    if (Trim(FolderPage.Values[0]) = '') or (not DirExists(FolderPage.Values[0])) then
    begin
      MsgBox('Choose a valid customer files folder.', mbError, MB_OK);
      Result := False;
      exit;
    end;
  end;

  if CurPageID = YuTemplatePage.ID then
  begin
    if (Trim(YuTemplatePage.Values[0]) = '') or (not DirExists(YuTemplatePage.Values[0])) then
    begin
      MsgBox('Choose a valid YU template folder.', mbError, MB_OK);
      Result := False;
      exit;
    end;

    if not FileExists(AddBackslash(Trim(YuTemplatePage.Values[0])) + 'yuchang_order_form_Widget.xlsx') then
    begin
      MsgBox(
        'The selected folder does not contain yuchang_order_form_Widget.xlsx.' + #13#10#13#10 +
        'Choose the folder that contains that workbook.',
        mbError,
        MB_OK);
      Result := False;
      exit;
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ConfigDir, ConfigPath, JsonText: string;
begin
  if CurStep = ssPostInstall then
  begin
    ConfigDir := ExpandConstant('{commonappdata}\WindsorWidget');
    if not DirExists(ConfigDir) then
      ForceDirectories(ConfigDir);

    ConfigPath := ConfigDir + '\client_config.json';
    JsonText :=
      '{' + #13#10 +
      '  "provider": "sqlserver",' + #13#10 +
      '  "driver": "ODBC Driver 18 for SQL Server",' + #13#10 +
      '  "server": "' + JsonEscape(Trim(SqlPage.Values[0])) + '",' + #13#10 +
      '  "port": "' + JsonEscape(Trim(SqlPage.Values[1])) + '",' + #13#10 +
      '  "database": "' + JsonEscape(Trim(SqlPage.Values[2])) + '",' + #13#10 +
      '  "username": "' + JsonEscape(Trim(SqlPage.Values[3])) + '",' + #13#10 +
      '  "password": "' + JsonEscape(Trim(SqlPage.Values[4])) + '",' + #13#10 +
      '  "trusted_connection": false,' + #13#10 +
      '  "encrypt": false,' + #13#10 +
      '  "trust_server_certificate": true,' + #13#10 +
      '  "timeout": 5' + #13#10 +
      '}' + #13#10;

    SaveStringToFile(ConfigPath, JsonText, False);
    RegWriteStringValue(HKCU, 'Software\Windsor\WidgetApp', 'customerFilesRoot', Trim(FolderPage.Values[0]));
    RegWriteStringValue(
      HKCU,
      'Software\Windsor\WidgetApp\yu',
      'template_path',
      AddBackslash(Trim(YuTemplatePage.Values[0])) + 'yuchang_order_form_Widget.xlsx'
    );
  end;
end;
