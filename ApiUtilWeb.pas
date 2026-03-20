unit ApiUtilWeb;

interface
{$WARNINGS ON}
{$HINTS ON}
{$WARN UNIT_PLATFORM OFF}
{$WARN SYMBOL_PLATFORM OFF}

uses
  Windows, Messages, SysUtils, variants, Classes,
  System.StrUtils, System.Types, registry, ActiveX,
  Vcl.Graphics, Vcl.Forms, FireDAC.Comp.Client,
  System.RegularExpressions, Winsock;

type
  WinIsWow64 = function( Handle: THandle; var Iret: BOOL ): Windows.BOOL; stdcall;

  TOfficeVersion = (OfficeUnknown,Office97,Office98,Office2000,OfficeXP,Office2003,Office2007,Office2010,
                     Office2013,Office2016,Office365);

procedure SleepForXXXSecend(Secunda : Integer; WaitToProcessMessage : Boolean);
procedure SleepForXXXMiliSecend(MiliSecunda : Integer; WaitToProcessMessage : Boolean);
function  CheckStrBool(Val : String) : Boolean;

function  ZeroIntToStr(Count : integer; Val : integer) : string;
function  BoolToStrInt(Val : Boolean) : String;
function  GetWinTempDir : String;
function  GetPcComputerName : String;
function  GetLocalIp: string;
function  LocalIP: String;
function  GetFDacCatalogName(FDacConn : TFDConnection) : String;
function  GetFDacConnProperty(FDacConn : TFDConnection; const PropName : String) : String;
function  GetComputerNetName: String;
function  GetComputerNetNameByExe(FileName : String) : String;
function  GetComputerNameExString(ANameFormat: COMPUTER_NAME_FORMAT): String;
function  GetGUIDString : String;
function  ConvertHebColorToInt(HebColor : String) : Integer;
function  Convert_AsciiHeb2UniCode(const StrAnsi : string): Widestring;

function  IsWindows64: Boolean;
function  GetOfficeNo: TOfficeVersion;
function  IsOffice64: Boolean;
procedure GetBuildInfo(FileName : WideString; var V1, V2, V3, V4: word);
function  GetExeFileVersion(FileName : WideString) : String;

function  RemoveAllNonNumericChars(Str : string) : string;
function  RemoveNonNumericChars(Str : string; UseExtra : Boolean) : string;
function  RemoveNonReadebleChars(Str : string) : string;

function  RemoveLastEnterChar(SrcString : String) : String;
function  RemoveAllEnterChar(SrcString : String) : String;
function  RemoveEnterChar(SrcString : String) : String;
function  RemoveAllTabChar(SrcString : WideString) : WideString;
function  RemoveBackSlashChar(SrcString : String) : String;
function  RemoveSlashChar(SrcString : String) : String;
function  FixDirWithSlash(Dir : String) : String;
function  AddBackSlashChar(SrcString : String) : String;
function  GetUniqueNumber : Integer;

function  CleanSQLKeyValue(FileName :String) : String;
function  CleanDirFileName(Dir :String) : String;
function  CleanUnPrintableChar(CheckString :String) : String;

function  GetKeyString : String;
function  CheckIsraelIdNumber(Id : Integer) : Boolean;
function  CheckIsraelCompId(CompId : Integer) : Boolean;
function  IsMatch(const Input, Pattern: string): boolean;
function  IsValidEmailRegEx(const EmailAddress: string): boolean;
function  IsPasswordValid(const s: AnsiString): Boolean;
function  IsValidPhoneNumber(const APhoneNumber: string): Boolean;
function  GetRegionByPrefix(Prefix : integer) : String;

function  CutStringToWords(FullString : String; WordList : TStringList; DoSort : Boolean = True) : Boolean; OverLoad;
function  CutStringToWords(FullString : String; WordList : TStringList; Delimiter : Char; DoSort : Boolean = True) : Boolean; OverLoad;
function  RemoveDuplicateWords(SearchText : String): String;
function  AddMsgToEventLog(FMachine     : String;
                           FEventSource : String;
                           EventMsg     : String;
                           EventType    : Integer;
                           Category     : Integer;
                           MessageID    : Integer) : Boolean;

function ConvertFileToBase64_(FileName : String; var Base64Str : String) : Boolean;
function ConvertFileToBase64Mime_(FileName : String; var Base64StrMime : String) : Boolean;
function ConvertBase64ToFile_(Base64Str : String; FileName : String) : Boolean;
function ConvertBase64ToFileMime_(Base64StrMime : String; FileName : String) : Boolean;

function FileToBase64_(const FilePath: string): string;
function Base64ToFile_(const Base64String, FilePath: string) : Boolean;
function TColorToHTMLColor(Color: TColor): string;
function IsSQLCommandDanger(SQLCOMMAND : String) : Boolean;

implementation

Uses
  System.Json, System.NetEncoding,
  IdBaseComponent, IdComponent, IdIPWatch,  uFileUtilWeb, IdHTTP,
  Global, BinData;

function GetComputerNameExW(NameType: COMPUTER_NAME_FORMAT; lpBuffer: LPWSTR;
  var nSize: DWORD): BOOL; stdcall; external kernel32 name 'GetComputerNameExW';

function  RemoveAllNonNumericChars(Str : string) : string;
var
  CurChr          : integer;
  CheckStr        : String;
begin
  CheckStr := Trim(Str);
  Result := '';
  for CurChr := 1 to Length(CheckStr) do
  begin
    if CheckStr[CurChr] in ['0'..'9'] then
      Result := Result + CheckStr[CurChr];
  end;

end;

function RemoveNonNumericChars(Str : string; UseExtra : Boolean) : string;
const
  NUMERIC_CHARS1 : set of char =  ['0' .. '9', '.'];
  NUMERIC_CHARS2 : set of char =  ['-', '+'];
var
  CurChr          : integer;
  MinusSignFound  : Boolean;
  CheckStr        : String;
begin
  MinusSignFound  := False;

  CheckStr := Trim(Str);
  IF UseExtra Then
  begin
    if (Copy(CheckStr,1,1) = '-') Or
       (Copy(CheckStr,Length(CheckStr),1) = '-') Then
      MinusSignFound  := True;
  end;

  Result := '';
  for CurChr := 1 to Length(CheckStr) do
  begin
    if CheckStr[CurChr] in NUMERIC_CHARS1 then
      Result := Result + CheckStr[CurChr];
  end;

  IF UseExtra Then
  begin
    IF MinusSignFound Then
      Result := '-' + Result;
  end;
end;

function  RemoveNonReadebleChars(Str : string) : string;
var
  CheckStr  : String;
  CurChar   : integer;
  CheckChar : Char;
begin
  CheckStr := Trim(Str);

  Result := '';
  for CurChar := 1 to Length(CheckStr) do
  begin
    CheckChar := CheckStr[CurChar];
    IF (CheckChar = #10) Then {LF}
       CheckChar := ' ';
    IF (CheckChar = #13) Then {CR}
       CheckChar := ' ';
    IF (CheckChar = #224) Then {א}
       CheckChar := 'א';
    IF (CheckChar = #225) Then {ב}
       CheckChar := 'ב';
    IF (CheckChar = #226) Then {ג}
       CheckChar := 'ג';
    IF (CheckChar = #227) Then {ד}
       CheckChar := 'ד';
    IF (CheckChar = #228) Then {ה}
       CheckChar := 'ה';
    IF (CheckChar = #229) Then {ו}
       CheckChar := 'ו';
    IF (CheckChar = #230) Then {ז}
       CheckChar := 'ז';
    IF (CheckChar = #231) Then {ח}
       CheckChar := 'ח';
    IF (CheckChar = #232) Then {ט}
       CheckChar := 'ט';
    IF (CheckChar = #233) Then {י}
       CheckChar := 'י';
    IF (CheckChar = #234) Then {כ}
       CheckChar := 'ך';
    IF (CheckChar = #235) Then {ך}
       CheckChar := 'כ';
    IF (CheckChar = #236) Then {ל}
       CheckChar := 'ל';
    IF (CheckChar = #237) Then {מ}
       CheckChar := 'ם';
    IF (CheckChar = #238) Then {ם}
       CheckChar := 'מ';
    IF (CheckChar = #239) Then {נ}
       CheckChar := 'ן';
    IF (CheckChar = #240) Then {ן}
       CheckChar := 'נ';
    IF (CheckChar = #241) Then {ס}
       CheckChar := 'ס';
    IF (CheckChar = #242) Then {ע}
       CheckChar := 'ע';
    IF (CheckChar = #243) Then {פ}
       CheckChar := 'ף';
    IF (CheckChar = #244) Then {ף}
       CheckChar := 'פ';
    IF (CheckChar = #245) Then {צ}
       CheckChar := 'ץ';
    IF (CheckChar = #246) Then {ץ}
       CheckChar := 'צ';
    IF (CheckChar = #247) Then {ק}
       CheckChar := 'ק';
    IF (CheckChar = #248) Then {ר}
       CheckChar := 'ר';
    IF (CheckChar = #249) Then {ש}
       CheckChar := 'ש';
    IF (CheckChar = #250) Then {ת}
       CheckChar := 'ת';

    if (CheckChar = '0') Or
       (CheckChar = '1') Or
       (CheckChar = '2') Or
       (CheckChar = '3') Or
       (CheckChar = '4') Or
       (CheckChar = '5') Or
       (CheckChar = '6') Or
       (CheckChar = '7') Or
       (CheckChar = '8') Or
       (CheckChar = '9') Or
       (CheckChar = #128) Or {א}
       (CheckChar = #129) Or {ב}
       (CheckChar = #130) Or {ג}
       (CheckChar = #131) Or {ד}
       (CheckChar = #132) Or {ה}
       (CheckChar = #133) Or {ו}
       (CheckChar = #134) Or {ז}
       (CheckChar = #135) Or {ח}
       (CheckChar = #136) Or {ט}
       (CheckChar = #137) Or {י}
       (CheckChar = #138) Or {כ}
       (CheckChar = #139) Or {ך}
       (CheckChar = #140) Or {ל}
       (CheckChar = #141) Or {מ}
       (CheckChar = #142) Or {ם}
       (CheckChar = #143) Or {נ}
       (CheckChar = #144) Or {ן}
       (CheckChar = #145) Or {ס}
       (CheckChar = #146) Or {ע}
       (CheckChar = #147) Or {פ}
       (CheckChar = #148) Or {ף}
       (CheckChar = #149) Or {צ}
       (CheckChar = #150) Or {ץ}
       (CheckChar = #151) Or {ק}
       (CheckChar = #152) Or {ר}
       (CheckChar = #153) Or {ש}
       (CheckChar = #154) Or {ת}
       (UpperCase(CheckChar) = UpperCase('A')) Or
       (UpperCase(CheckChar) = UpperCase('B')) Or
       (UpperCase(CheckChar) = UpperCase('C')) Or
       (UpperCase(CheckChar) = UpperCase('D')) Or
       (UpperCase(CheckChar) = UpperCase('E')) Or
       (UpperCase(CheckChar) = UpperCase('F')) Or
       (UpperCase(CheckChar) = UpperCase('G')) Or
       (UpperCase(CheckChar) = UpperCase('H')) Or
       (UpperCase(CheckChar) = UpperCase('I')) Or
       (UpperCase(CheckChar) = UpperCase('J')) Or
       (UpperCase(CheckChar) = UpperCase('K')) Or
       (UpperCase(CheckChar) = UpperCase('L')) Or
       (UpperCase(CheckChar) = UpperCase('M')) Or
       (UpperCase(CheckChar) = UpperCase('N')) Or
       (UpperCase(CheckChar) = UpperCase('O')) Or
       (UpperCase(CheckChar) = UpperCase('P')) Or
       (UpperCase(CheckChar) = UpperCase('Q')) Or
       (UpperCase(CheckChar) = UpperCase('R')) Or
       (UpperCase(CheckChar) = UpperCase('S')) Or
       (UpperCase(CheckChar) = UpperCase('T')) Or
       (UpperCase(CheckChar) = UpperCase('U')) Or
       (UpperCase(CheckChar) = UpperCase('V')) Or
       (UpperCase(CheckChar) = UpperCase('W')) Or
       (UpperCase(CheckChar) = UpperCase('X')) Or
       (UpperCase(CheckChar) = UpperCase('Y')) Or
       (UpperCase(CheckChar) = UpperCase('Z')) Or
       (CheckChar = 'א') Or
       (CheckChar = 'ב') Or
       (CheckChar = 'ג') Or
       (CheckChar = 'ד') Or
       (CheckChar = 'ה') Or
       (CheckChar = 'ו') Or
       (CheckChar = 'ז') Or
       (CheckChar = 'ח') Or
       (CheckChar = 'ט') Or
       (CheckChar = 'י') Or
       (CheckChar = 'כ') Or
       (CheckChar = 'ך') Or
       (CheckChar = 'ל') Or
       (CheckChar = 'מ') Or
       (CheckChar = 'ם') Or
       (CheckChar = 'נ') Or
       (CheckChar = 'ן') Or
       (CheckChar = 'ס') Or
       (CheckChar = 'ע') Or
       (CheckChar = 'פ') Or
       (CheckChar = 'ף') Or
       (CheckChar = 'צ') Or
       (CheckChar = 'ץ') Or
       (CheckChar = 'ק') Or
       (CheckChar = 'ר') Or
       (CheckChar = 'ש') Or
       (CheckChar = 'ת') Or
       (CheckChar = '_') Or
       (CheckChar = '-') Or
       (CheckChar = '~') Or
       (CheckChar = '`') Or
       (CheckChar = ';') Or
       (CheckChar = '!') Or
       (CheckChar = '@') Or
       (CheckChar = '#') Or
       (CheckChar = '$') Or
       (CheckChar = '%') Or
       (CheckChar = '^') Or
       (CheckChar = '&') Or
       (CheckChar = '*') Or
       (CheckChar = ')') Or
       (CheckChar = '(') Or
       (CheckChar = '=') Or
       (CheckChar = '+') Or
       (CheckChar = ']') Or
       (CheckChar = '[') Or
       (CheckChar = '}') Or
       (CheckChar = '{') Or
       (CheckChar = ':') Or
       (CheckChar = '"') Or
       (CheckChar = '''') Or
       (CheckChar = '|') Or
       (CheckChar = '\') Or
       (CheckChar = '/') Or
       (CheckChar = '?') Or
       (CheckChar = '.') Or
       (CheckChar = '>') Or
       (CheckChar = '<') Or
       (CheckChar = ',') Or
       (CheckChar = #32) Or
       (CheckChar = ' ') then
    begin
      Result := Result + CheckChar;
    end;
  end;
end;

function  RemoveLastEnterChar(SrcString : String) : String;
begin
  // Remove The Last $D$A From SrcString;

  Result := SrcString;
  While True Do
  begin
    IF (Copy(Result,Length(Result),1) = #13) Or
       (Copy(Result,Length(Result),1) = #10) Then
    begin
      Result := Copy(Result,1,Length(Result)-1);
    end
    else
      Break;
  end;
end;

procedure SleepForXXXSecend(Secunda : Integer; WaitToProcessMessage : Boolean);
Var
  CurTime : Integer;
Const
  LoopFor : Integer = 50;
begin
  For CurTime := 1 to Secunda Do
  begin
    Sleep(250);
    if WaitToProcessMessage then
      Application.ProcessMessages;
    Sleep(250);
    if WaitToProcessMessage then
      Application.ProcessMessages;
    Sleep(250);
    if WaitToProcessMessage then
      Application.ProcessMessages;
    Sleep(250);
    if WaitToProcessMessage then
      Application.ProcessMessages;
  end;
end;

procedure SleepForXXXMiliSecend(MiliSecunda : Integer; WaitToProcessMessage : Boolean);
Var
  CurTime  : Integer;
  UntilMax : Integer;
begin
  UntilMax := Trunc(MiliSecunda / 10);
  if UntilMax < 5 then
    UntilMax := 5;
  For CurTime := 1 to UntilMax Do
  begin
    Sleep(09);  {10 cose it work more then excepted}
    if WaitToProcessMessage then
      Application.ProcessMessages;
  end;
end;

Function GetSystemWow64Directory (lpBuffer: LPTSTR; uSize: Uint): uint; stdcall; external 'Kernel32.dll' name 'GetSystemWow64DirectoryW';

Function GetSys64Dir: string;
var
  Buffer: Array [0..max_path] of Char;
begin
  Result := '';
  If GetSystemWow64Directory (Buffer, max_path) = 0 Then
    Exit;
  Result := IncludeTrailingBackslash(string (Buffer));
end;

function  CheckStrBool(Val : String) : Boolean;
Begin
   Result := (Val = '1') Or (UpperCase(Val) = UpperCase('t'));
End;

function ZeroIntToStr(Count : integer; Val : integer) : string;
Var
  Str : String;
  Lng : Integer;
begin
  //Result := Copy(Format('%.' + IntToStr(Count) + 'd', [Val]), 1, Count);
  Try
    Str := IntToStr(Val);
    Lng := Length(Str);
    IF Count > Lng Then
      Str := StringOfChar('0', Count - Lng) + Str;
    Result := Copy(Str,1,Count);
  Except;
    Try
      Result  := IntToStr(Val);
      While Length(Result) < Count Do
        Result  := '0' + Result;
    Except;
      Result := '';
    end;
  end;
end;

function  BoolToStrInt(Val : Boolean) : String;
begin
  Result := '0';
  if Val then
    Result := '1';
end;

function GetWinTempDir : String;
var
  le:Integer;
Const
  MAX_PATH : Dword = 1024;
begin
  SetLength(result, MAX_PATH);
  le:=GetTempPath(MAX_PATH, PChar(result));
  SetLength(result,le);
end;

function  GetPcComputerName : String;
var
  Pc: PChar;
  Sz: DWORD;
begin
  Result := '';
  Try
    Sz := MAX_COMPUTERNAME_LENGTH+1;
    Pc := strAlloc(Sz);
    Try
      // Computer name
      if getComputerName(Pc, Sz) then
        Result := Pc;
    Finally
      strDispose(Pc);
    End;
  Except;
    Result := 'Exception';
  End;
end;

function GetLocalIp: string;
var
   IPW: TIdIPWatch;
begin
  Result := '';
  IpW := TIdIPWatch.Create(nil);
  try
    if IpW.LocalIP <> '' then
      Result := IpW.LocalIP;
  finally
    IpW.Free;
  end;
end;

function LocalIP: String;
type
  TaPInAddr = array[0..10] of PInAddr;
  PaPInAddr = ^TaPInAddr;
var
  phe: PHostEnt;
  pptr: PaPInAddr;
  Buffer: array[0..63] of AnsiChar;
  I: Integer;
  GInitData: TWSAData;
begin
  WSAStartup($101, GInitData);
  Try
    Result := '';
    GetHostName(Buffer, SizeOf(Buffer));
    phe := GetHostByName(buffer);
    if phe = nil then Exit;
    pPtr := PaPInAddr(phe^.h_addr_list);
    I := 0;
    while pPtr^[I] <> nil do
    begin
      Result := inet_ntoa(pptr^[I]^);
      Inc(I);
    end;
  Finally
    WSACleanup;
  End;
end;

function  GetFDacCatalogName(FDacConn : TFDConnection) : String;
Var
  CurItem : integer;
  lConnName : String;
  StrVal  : String;
  ACnnlist : TStringList;
  Alist : TStringList;
begin
  Result := '';
  if not FDacConn.Connected then
    Exit;

  Try
    ACnnlist := TStringList.Create;
    Alist    := TStringList.Create;
    //FDManager.GetConnectionNames(ACnnlist);
    //For CurCnn := 0 To ACnnlist.Count - 1 Do
    //begin
      //lConnName := ACnnlist.Strings[CurCnn];
      lConnName := FDacConn.Name;
      //FDManager.GetConnectionDefParams(lConnName,Alist);
      FDacConn.GetCatalogNames('',Alist);
      For CurItem := 0 To Alist.Count - 1 Do
      begin
        StrVal  := UpperCase(Alist.Strings[CurItem]);
        if StrVal <> '' Then
        begin
          Result := Trim(StrVal);
          Break;
        end;
      end;
  Finally
    Alist.Free;
    ACnnlist.Free;
  End;

end;

function  GetFDacConnProperty(FDacConn : TFDConnection; const PropName : String) : String;
Var
  i : integer;
  CurItem : integer;
  lConnName : String;
  StrVal  : String;
  P_Name  : String;
  P_Value : String;
  ACnnlist : TStringList;
  Alist : TStringList;
begin
(*
ConnectionDef=MyCONN
DriverID=MSSQL
Server=hanibaal\sqlexpress
Database=mzcatalog
User_Name=sa
Password=a334.....
MetaDefSchema=dbo
MetaDefCatalog=mzcatalog
ExtendedMetadata=True
OSAuthent=No
Integrated Security=False
Name=MyCONN*)
  Result := '';
  if not FDacConn.Connected then
    Exit;
(*
  for I := 0 to AdoConn.Properties.Count - 1 do
  begin
    P_Name  := UpperCase(AdoConn.Properties.Item[I].Name);
    P_Value := UpperCase(VarToWideStr(AdoConn.Properties.Item[I].Value));
    if UpperCase(P_Name) = UpperCase(PropName) then
    begin
      Result := P_Value;
      Break;
    end;
  end;
*)

  Try
    ACnnlist := TStringList.Create;
    Alist    := TStringList.Create;
    //FDManager.GetConnectionNames(ACnnlist);
    //For CurCnn := 0 To ACnnlist.Count - 1 Do
    //begin
      //lConnName := ACnnlist.Strings[CurCnn];
      lConnName := FDacConn.ConnectionDefName;
      //FDManager.GetConnectionDefParams(lConnName,Alist);
      FDManager.GetConnectionDefNames(Alist);
      For CurItem := 0 To Alist.Count - 1 Do
      begin
        StrVal  := UpperCase(Alist.Strings[CurItem]);
        if Copy(StrVal,1,Length(PropName)) = UpperCase(PropName) then
        begin
          P_Value := Copy(StrVal,Length(PropName)+1,Length(StrVal));
          for i := 1 To Length(P_Value) Do
          begin
            if Copy(P_Value,i,1) = '=' then
            begin
              P_Value[i] := ' ';
              Break;
            end;
          end;
          Result := Trim(P_Value);
          Break;
        end;
      end;
    //end;
  Finally
    Alist.Free;
    ACnnlist.Free;
  End;

  if Result = '' then
  begin
    StrVal := '';
    For CurItem := 0 To FDacConn.Params.Count - 1 Do
    begin
      StrVal := UpperCase(FDacConn.Params.Strings[CurItem]);
      P_Name := Copy(StrVal,1,Length(PropName));
      if UpperCase(P_Name) = UpperCase(PropName) then
      begin
        P_Value := Copy(StrVal,Length(PropName)+1,Length(StrVal));
        for i := 1 To Length(P_Value) Do
        begin
          if Copy(P_Value,i,1) = '=' then
          begin
            P_Value[i] := ' ';
            Break;
          end;
        end;
        P_Value := Trim(P_Value);

        if UpperCase(P_Name) = UpperCase('ConnectionDef') Then
          Result := P_Value
        else
        if UpperCase(P_Name) = UpperCase('DriverID') Then
          Result := P_Value {FDacConn.Params.DriverID}
        else
        if UpperCase(P_Name) = UpperCase('Server') Then
          Result := P_Value
        else
        if UpperCase(P_Name) = UpperCase('Database') Then
          Result := P_Value {FDacConn.Params.Database}
        else
        if UpperCase(P_Name) = UpperCase('User_Name') Then
          Result := P_Value {FDacConn.Params.UserName}
        else
        if UpperCase(P_Name) = UpperCase('Password') Then
          Result := P_Value {FDacConn.Params.Password}
        else
        if UpperCase(P_Name) = UpperCase('MetaDefSchema') Then
          Result := P_Value
        else
        if UpperCase(P_Name) = UpperCase('MetaDefCatalog') Then
          Result := P_Value
        else
        if UpperCase(P_Name) = UpperCase('ExtendedMetadata') Then
          Result := P_Value
        else
        if UpperCase(P_Name) = UpperCase('OSAuthent') Then
          Result := P_Value
        else
        if UpperCase(P_Name) = UpperCase('Integrated Security') Then
          Result := P_Value
        else
        if UpperCase(P_Name) = UpperCase('Name') Then
          Result := P_Value;

         if Result <> '' then
          Break;
      end;
    end;
  end;
end;

//******************************************************************************
//Call the WindowsAPI function to get the name of the computer from which this
//call was made.
//This just wraps the API function in a way that returns standard String result.
//Note: you will need the "Windows" unit in your USES clause.
//
//   In What Name Is Your Computer Is Called In The Net
//******************************************************************************
function GetComputerNetName: String;
var
  NameBuffer: array[0..255] of char;
  NameString : String;
  size: dword;
begin
  Result := '';

  try
    size := 256;
    if GetComputerName(NameBuffer, size) then
    begin
      NameString := NameBuffer;
      Result := Trim(NameString);
    end;
  except
    // eat errors
  end;
end; //GetComputerNetName

function  GetComputerNetNameByExe(FileName : String) : String;
Var
  Str       : String;
  CurChar   : Integer;
  FromStart : Integer;
begin
  Str       := ExpandUNCFileName(FileName);
  FromStart := 1;
  Result    := '';
  IF Copy(Str,1,2) = '\\' Then
  begin
    FromStart := 3;
    Result := '\\';
  end;

  For CurChar := FromStart To Length(Str) Do
  begin
    IF Str[CurChar] = '\' Then
      Break;
    Result := Result + Str[CurChar];
  end;
end;

function GetComputerNameExString(ANameFormat: COMPUTER_NAME_FORMAT): String;
var
  nSize: DWORD;
begin
  // to get domain name use
  // DomainName := GetComputerNameExString(ComputerNameDnsDomain)
  nSize := 1024;
  SetLength(Result, nSize);
  if GetComputerNameExW(ANameFormat, PWideChar(Result), nSize) then
    SetLength(Result, nSize)
  else
    Result := '';
  Result := Trim(Result);
end;

function GetGUIDString : String;
var
  udtGUID : TGUID;
  lResult : longint;
begin
{
Create a GUID at runtime
You may need to create a GUID at runtime. One reason could be to simply have a unique number that identifies each workstation.
By the way the last group of digits in the GUID generated on a computer is its network card number (if any is present).
Delphi does not have the needed definitions, here they are:

Use OLE32 DLL Decleration
function CoCreateGuid (pGUID : TGUID) : longint;
  external 'OLE32.DLL';

}

  Result := '';

  lResult := CoCreateGuid(udtGUID);
  // see definition of TGUID in Delphi's online help
  // udtGUID.D4 = network card's number
  If lResult = S_OK Then
  Result := GUIDToString(udtGUID);
end;

function IsWindows64: Boolean;
Var
  Directory : Array [0..max_path] of Char;
begin
  Result := False;
  if (GetSystemWow64Directory(Directory, Max_Path) > 0) Then
     // OS is 64bit
    Result := True;
end;

function  ConvertHebColorToInt(HebColor : String) : Integer;
begin
  Result := 0;
  if HebColor = 'כחול' Then
  begin
    Result := 16724736;
  end
  else
  if HebColor = 'אדום' Then
  begin
    Result := 255;
  end
  else
  if HebColor = 'ירוק' Then
  begin
    Result := 65331;
  end
  else
  if HebColor = 'צהוב' Then
  begin
    Result := 65535;
  end
  else
  if HebColor = 'תכלת' Then
  begin
    Result := 16777062;
  end
  else
  if HebColor = 'זית' Then
  begin
    Result := 39270;
  end
  else
  if HebColor = 'כסף' Then
  begin
    Result := 13421772;
  end
  else
  if HebColor = 'לימון' Then
  begin
    Result := 65280;
  end
  else
  if HebColor = 'כחול ים' Then
  begin
    Result := 16711782;
  end
  else
  if HebColor = 'סגול' Then
  begin
    Result := 10027161;
  end
  else
  if HebColor = 'אדום עז' Then
  begin
    Result := 13209;
  end
  else
  if HebColor = 'ורוד' Then
  begin
    Result := 16751103;
  end
  else
  if HebColor = 'טורקיז' Then
  begin
    Result := 13421568;
  end;
end;

function Convert_AsciiHeb2UniCode(const StrAnsi : string): Widestring;
var
  strumien : TStringStream;
  reader   : TStreamReader;
begin
  Result := '';
  strumien := TStringStream.Create(StrAnsi);
  Try
    reader:=TStreamReader.Create(strumien, TEncoding.GetEncoding(862));
    Try
      reader.OwnStream;
      while not reader.EndOfStream do
        Result := Result + reader.ReadLine;
    Finally
      reader.Free;
    End;
  Finally
    strumien := nil;
  End;
end;

function  GetOfficeNo: TOfficeVersion;
Var
  Reg    : TRegistry;
  RegKey : String;
  OfficeVesion : String;
Const
  PDRoot = HKEY_CLASSES_ROOT;
begin
  Result := OfficeUnknown;

  Reg := TRegistry.Create;
  Try
    Reg.RootKey := PDRoot;
    RegKey := 'Word.Application\CurVer';
    If Reg.OpenKeyReadOnly(RegKey) Then
    begin
      OfficeVesion := RemoveAllNonNumericChars(Reg.ReadString(''));
      Reg.CloseKey;
    end;

    //Office 97   -  7.0
    //Office 98   -  8.0
    //Office 2000 -  9.0
    //Office XP   - 10.0
    //Office 2003 - 11.0
    //Office 2007 - 12.0
    //Office 2010 - 14.0 (sic!)
    //Office 2013 - 15.0
    //Office 2016 - 16.0

    If (OfficeVesion = '7') or (OfficeVesion = '07') or (OfficeVesion = '7.0') Then
      Result := Office97
    else
    If (OfficeVesion = '8') or (OfficeVesion = '08') or (OfficeVesion = '8.0') Then
      Result := Office98
    else
    If (OfficeVesion = '9') or (OfficeVesion = '09') or (OfficeVesion = '9.0') Then
      Result := Office2000
    else
    If (OfficeVesion = '10') or (OfficeVesion = '10.0') Then
      Result := OfficeXP
    else
    If (OfficeVesion = '11') or (OfficeVesion = '11.0') Then
      Result := Office2003
    else
    If (OfficeVesion = '12') or (OfficeVesion = '12.0') Then
      Result := Office2007
    else
    If (OfficeVesion = '14') or (OfficeVesion = '14.0') Then
      Result := Office2010
    else
    If (OfficeVesion = '15') or (OfficeVesion = '15.0') Then
      Result := Office2013
    else
    If (OfficeVesion = '16') or (OfficeVesion = '16.0') Then
      Result := Office2016;

  Finally
    Reg.Free;
  End;
end;

function  IsOffice64: Boolean;
(*
Remember that the 64-bit version of Office can be installed on 64-bit Windows only.

If Outlook is installed, then the value below exists in this registry key:

Outlook 2010-2016:
Registry view: both 32-bit and 64-bit
Key: HKLM\SOFTWARE\Microsoft\Office\{14, 15 or 16}.0\Outlook\InstallRoot
Value name: Bitness
That value can be "x64" or "x86"; "x64" means Outlook 64-bit is installed.

If Outlook is not installed, you can check the following values in the following 64-bit registry key:

Excel, Word, PowerPoint 2010-2016:
Registry view: 64-bit
Key: HKLM\Microsoft\Office\{14, 15 or 16}.0\{application}\InstallRoot
Value name: Path
If that value exists, then the corresponding 64-bit application is installed.
*)

Var
  Reg    : TRegistry;
  RegKey : String;
  CurVer : Integer;
Const
  PDRoot = HKEY_LOCAL_MACHINE;
  PDPath1  : String = 'Software\Microsoft\Office\XX.0\Outlook';
  PDPath2  : String = 'Software\Microsoft\Office\XX.0\Word';
  PDPath3  : String = 'SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\XX.0\Outlook';
  PDPath4  : String = 'SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\XX.0\word';
  PDPath5  : String = 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration';
  PDPath6  : String = 'SOFTWARE\WOW6432Node\Microsoft\Office\XX.0\Outlook';
  PDPath7  : String = 'SOFTWARE\WOW6432Node\Microsoft\Office\XX.0\word';
  PDPath8  : String = 'SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\XX.0\Outlook';
  PDPath9  : String = 'SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\XX.0\Word';
  PDPath10 : String = 'SOFTWARE\Microsoft\Office\XX.0\Outlook\InstallRoot';

type
  TRegistryView = (rvDefault, rvRegistry64, rvRegistry32);

  function RegistryViewAccessFlag(View: TRegistryView): LongWord;
  begin
    case View of
    rvDefault:
      Result := 0;
    rvRegistry64:
      Result := KEY_WOW64_64KEY;
    rvRegistry32:
      Result := KEY_WOW64_32KEY;
    end;
  end;

  function IsBitnessExists(VerRegKey : String) : Boolean;
  Var
    CheckRegKey : String;
  begin
    Result := False;
    CheckRegKey := StringReplace(PDPath1,'XX',VerRegKey,[rfReplaceAll, rfIgnoreCase]);
    if Reg.OpenKeyReadOnly(CheckRegKey) then
    begin
      if Reg.ValueExists('Bitness') then
        Result := UpperCase(Reg.ReadString('Bitness')) = UpperCase('X64');
      Reg.CloseKey;
    end;

    if Not Result then
    begin
      CheckRegKey := StringReplace(PDPath5,'XX',VerRegKey,[rfReplaceAll, rfIgnoreCase]);
      if Reg.OpenKeyReadOnly(CheckRegKey) then
      begin
        if Reg.ValueExists('Platform') then
          Result := UpperCase(Reg.ReadString('Platform')) = UpperCase('X64');
        Reg.CloseKey;
      end;
    end;

    if Not Result then
    begin
      CheckRegKey := StringReplace(PDPath3,'XX',VerRegKey,[rfReplaceAll, rfIgnoreCase]);
      if Reg.OpenKeyReadOnly(CheckRegKey) then
      begin
        if Reg.ValueExists('Bitness') then
          Result := UpperCase(Reg.ReadString('Bitness')) = UpperCase('X64');
        Reg.CloseKey;
      end;
    end;

    if Not Result then
    begin
      CheckRegKey := StringReplace(PDPath6,'XX',VerRegKey,[rfReplaceAll, rfIgnoreCase]);
      if Reg.OpenKeyReadOnly(CheckRegKey) then
      begin
        if Reg.ValueExists('Bitness') then
          Result := UpperCase(Reg.ReadString('Bitness')) = UpperCase('X64');
        Reg.CloseKey;
      end;
    end;

    if Not Result then
    begin
      CheckRegKey := StringReplace(PDPath8,'XX',VerRegKey,[rfReplaceAll, rfIgnoreCase]);
      if Reg.OpenKeyReadOnly(CheckRegKey) then
      begin
        if Reg.ValueExists('Bitness') then
          Result := UpperCase(Reg.ReadString('Bitness')) = UpperCase('X64');
        Reg.CloseKey;
      end;
    end;

    if Not Result then
    begin
      CheckRegKey := StringReplace(PDPath10,'XX',VerRegKey,[rfReplaceAll, rfIgnoreCase]);
      if Reg.OpenKeyReadOnly(CheckRegKey) then
      begin
        if Reg.ValueExists('Bitness') then
          Result := UpperCase(Reg.ReadString('Bitness')) = UpperCase('X64');
        Reg.CloseKey;
      end;
    end;
  end;

  function IsWordBitnessExists(VerRegKey : String) : Boolean;
  Var
    CheckRegKey : String;
  begin
    Result := False;
    CheckRegKey := StringReplace(PDPath2,'XX',VerRegKey,[rfReplaceAll, rfIgnoreCase]);
    if Reg.OpenKeyReadOnly(CheckRegKey) then
    begin
      if Reg.ValueExists('Bitness') then
        Result := UpperCase(Reg.ReadString('Bitness')) = UpperCase('X64');
      Reg.CloseKey;
    end;

    if not Result then
    begin
      CheckRegKey := StringReplace(PDPath4,'XX',VerRegKey,[rfReplaceAll, rfIgnoreCase]);
      if Reg.OpenKeyReadOnly(CheckRegKey) then
      begin
        if Reg.ValueExists('Bitness') then
          Result := UpperCase(Reg.ReadString('Bitness')) = UpperCase('X64');
        Reg.CloseKey;
      end;
    end;

    if not Result then
    begin
      CheckRegKey := StringReplace(PDPath7,'XX',VerRegKey,[rfReplaceAll, rfIgnoreCase]);
      if Reg.OpenKeyReadOnly(CheckRegKey) then
      begin
        if Reg.ValueExists('Bitness') then
          Result := UpperCase(Reg.ReadString('Bitness')) = UpperCase('X64');
        Reg.CloseKey;
      end;
    end;

    if not Result then
    begin
      CheckRegKey := StringReplace(PDPath9,'XX',VerRegKey,[rfReplaceAll, rfIgnoreCase]);
      if Reg.OpenKeyReadOnly(CheckRegKey) then
      begin
        if Reg.ValueExists('Bitness') then
          Result := UpperCase(Reg.ReadString('Bitness')) = UpperCase('X64');
        Reg.CloseKey;
      end;
    end;
  end;

begin
  Result := False;

//=======================================================================
  Reg := TRegistry.Create(KEY_READ or RegistryViewAccessFlag(rvDefault));
  Try
    Reg.RootKey := PDRoot;

    For CurVer := 14 To 25 Do
    begin
      Result := IsBitnessExists(IntToStr(CurVer));
      if Result then
        Break;
    end;

    if Not Result Then
    begin
      For CurVer := 14 To 25 Do
      begin
        Result := IsWordBitnessExists(IntToStr(CurVer));
        if Result then
          Break;
      end;
    end;
  Finally
    Reg.Free;
  End;

  if Result then
    Exit;

//=======================================================================
  Reg := TRegistry.Create(KEY_READ or RegistryViewAccessFlag(rvRegistry32));
  Try
    Reg.RootKey := PDRoot;

    For CurVer := 14 To 25 Do
    begin
      Result := IsBitnessExists(IntToStr(CurVer));
      if Result then
        Break;
    end;

    if Not Result Then
    begin
      For CurVer := 14 To 25 Do
      begin
        Result := IsWordBitnessExists(IntToStr(CurVer));
        if Result then
          Break;
      end;
    end;
  Finally
    Reg.Free;
  End;

  if Result then
    Exit;

//=======================================================================
  Reg := TRegistry.Create(KEY_READ or RegistryViewAccessFlag(rvRegistry64));
  Try
    Reg.RootKey := PDRoot;

    For CurVer := 14 To 25 Do
    begin
      Result := IsBitnessExists(IntToStr(CurVer));
      if Result then
        Break;
    end;

    if Not Result Then
    begin
      For CurVer := 14 To 25 Do
      begin
        Result := IsWordBitnessExists(IntToStr(CurVer));
        if Result then
          Break;
      end;
    end;
  Finally
    Reg.Free;
  End;

end;

procedure GetBuildInfo(FileName : WideString; var V1, V2, V3, V4: word);
var
  VerInfoSize, VerValueSize, Dummy: DWORD;
  VerInfo: Pointer;
  VerValue: PVSFixedFileInfo;
begin
  VerInfoSize := GetFileVersionInfoSize(PChar(FileName), Dummy);
  if VerInfoSize > 0 then
  begin
      GetMem(VerInfo, VerInfoSize);
      try
        if GetFileVersionInfo(PChar(FileName), 0, VerInfoSize, VerInfo) then
        begin
          VerQueryValue(VerInfo, '\', Pointer(VerValue), VerValueSize);
          with VerValue^ do
          begin
            V1 := dwFileVersionMS shr 16;
            V2 := dwFileVersionMS and $FFFF;
            V3 := dwFileVersionLS shr 16;
            V4 := dwFileVersionLS and $FFFF;
          end;
        end;
      finally
        FreeMem(VerInfo, VerInfoSize);
      end;
  end;
end;

function GetExeFileVersion(FileName : WideString) : String;
var
  V1, V2, V3, V4: word;
begin
  GetBuildInfo(FileName, V1, V2, V3, V4);
  Result := IntToStr(V1) + '.' + IntToStr(V2) + '.' +
    IntToStr(V3) + '.' + IntToStr(V4);
end;

function  RemoveAllEnterChar(SrcString : String) : String;
begin
  // Remove The Last $D$A From SrcString;
  Result := SrcString;
  Result := StringReplace(Result,#13+#10,' ',[rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result,#13,' ',[rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result,#10,' ',[rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result,#$D#$A,' ',[rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result,#$D,' ',[rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result,#$A,' ',[rfReplaceAll, rfIgnoreCase]);
end;

function  RemoveEnterChar(SrcString : String) : String;
begin
  // Remove The Last $D$A From SrcString;

  Result := SrcString;
  While True Do
  begin
    IF (Copy(Result,Length(Result),1) = #13) Or
       (Copy(Result,Length(Result),1) = #10) Then
    begin
      Result := Copy(Result,1,Length(Result)-1);
    end
    else
      Break;
  end;
end;

function  RemoveAllTabChar(SrcString : WideString) : WideString;
begin
  // Remove The Last $D$A From SrcString;
  Result := SrcString;
  Result := StringReplace(Result,#09,' ',[rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result,#9,' ',[rfReplaceAll, rfIgnoreCase]);
end;

function  RemoveBackSlashChar(SrcString : String) : String;
begin
  // Remove The Last '\' From SrcString;
  Result := RemoveLastEnterChar(SrcString);
  While True Do
  begin
    IF Copy(Result,Length(Result),1) = '\' Then
    begin
      Result := Copy(Result,1,Length(Result)-1);
    end
    else
    begin
      Break;
    end;
  end;
end;

function  RemoveSlashChar(SrcString : String) : String;
begin
  // Remove The Last '\' From SrcString;
  Result := RemoveLastEnterChar(SrcString);
  While True Do
  begin
    IF Copy(Result,Length(Result),1) = '/' Then
    begin
      Result := Copy(Result,1,Length(Result)-1);
    end
    else
    begin
      Break;
    end;
  end;
end;

function  FixDirWithSlash(Dir : String) : String;
begin
  Result := RemoveBackSlashChar(Dir) + '\';
end;

function  AddBackSlashChar(SrcString : String) : String;
begin
  // Add '\' To End Of SrcString;
  Result := RemoveBackSlashChar(SrcString);
  Result := Result + '\';
end;

function  GetUniqueNumber : Integer;
Var
  FloatNo        : single;
  NewSeed_1_Int  : Int64;
  NewSeed_2_Int  : Int64;
  NewSeed_2_Str  : String;
  ResultInt      : Int64;
  ResultStr      : String;
  LastStrRslt    : String;
begin
  Randomize;
  NewSeed_1_Int := GetTickCount64;

  NewSeed_2_Str  := FormatDateTime('yyssnnzzzhh', Now);
  NewSeed_2_Int  := Abs(StrToInt64Def(NewSeed_2_Str, 97879887));

  FloatNo := Random;

  LastStrRslt := '123056789';
  ResultInt := Abs(Trunc(NewSeed_1_Int * NewSeed_2_Int * FloatNo));
  ResultStr := IntToStr(ResultInt);
  if Length(ResultStr) > 9 then
  begin
    LastStrRslt := Copy(ResultStr,1,9);
  end
  else
  begin
    LastStrRslt := ResultStr;
  end;

  Result := StrToInt(LastStrRslt);
end;

function  CleanSQLKeyValue(FileName :String) : String;
Var
  CurChar : Integer;
  CheckStr : String;
  CheckChar : Char;
begin
  Result := '';
  CheckStr := Trim(FileName);
  if CheckStr = '' then
    Exit;
  for CurChar := 1 to Length(CheckStr) do
  begin
    CheckChar := CheckStr[CurChar];

    IF (CheckChar = #224) Then {א}
       CheckChar := 'א';
    IF (CheckChar = #225) Then {ב}
       CheckChar := 'ב';
    IF (CheckChar = #226) Then {ג}
       CheckChar := 'ג';
    IF (CheckChar = #227) Then {ד}
       CheckChar := 'ד';
    IF (CheckChar = #228) Then {ה}
       CheckChar := 'ה';
    IF (CheckChar = #229) Then {ו}
       CheckChar := 'ו';
    IF (CheckChar = #230) Then {ז}
       CheckChar := 'ז';
    IF (CheckChar = #231) Then {ח}
       CheckChar := 'ח';
    IF (CheckChar = #232) Then {ט}
       CheckChar := 'ט';
    IF (CheckChar = #233) Then {י}
       CheckChar := 'י';
    IF (CheckChar = #234) Then {כ}
       CheckChar := 'ך';
    IF (CheckChar = #235) Then {ך}
       CheckChar := 'כ';
    IF (CheckChar = #236) Then {ל}
       CheckChar := 'ל';
    IF (CheckChar = #237) Then {מ}
       CheckChar := 'ם';
    IF (CheckChar = #238) Then {ם}
       CheckChar := 'מ';
    IF (CheckChar = #239) Then {נ}
       CheckChar := 'ן';
    IF (CheckChar = #240) Then {ן}
       CheckChar := 'נ';
    IF (CheckChar = #241) Then {ס}
       CheckChar := 'ס';
    IF (CheckChar = #242) Then {ע}
       CheckChar := 'ע';
    IF (CheckChar = #243) Then {פ}
       CheckChar := 'ף';
    IF (CheckChar = #244) Then {ף}
       CheckChar := 'פ';
    IF (CheckChar = #245) Then {צ}
       CheckChar := 'ץ';
    IF (CheckChar = #246) Then {ץ}
       CheckChar := 'צ';
    IF (CheckChar = #247) Then {ק}
       CheckChar := 'ק';
    IF (CheckChar = #248) Then {ר}
       CheckChar := 'ר';
    IF (CheckChar = #249) Then {ש}
       CheckChar := 'ש';
    IF (CheckChar = #250) Then {ת}
       CheckChar := 'ת';

    if (CheckChar = '0') Or
       (CheckChar = '1') Or
       (CheckChar = '2') Or
       (CheckChar = '3') Or
       (CheckChar = '4') Or
       (CheckChar = '5') Or
       (CheckChar = '6') Or
       (CheckChar = '7') Or
       (CheckChar = '8') Or
       (CheckChar = '9') Or
       (CheckChar = #128) Or {א}
       (CheckChar = #129) Or {ב}
       (CheckChar = #130) Or {ג}
       (CheckChar = #131) Or {ד}
       (CheckChar = #132) Or {ה}
       (CheckChar = #133) Or {ו}
       (CheckChar = #134) Or {ז}
       (CheckChar = #135) Or {ח}
       (CheckChar = #136) Or {ט}
       (CheckChar = #137) Or {י}
       (CheckChar = #138) Or {כ}
       (CheckChar = #139) Or {ך}
       (CheckChar = #140) Or {ל}
       (CheckChar = #141) Or {מ}
       (CheckChar = #142) Or {ם}
       (CheckChar = #143) Or {נ}
       (CheckChar = #144) Or {ן}
       (CheckChar = #145) Or {ס}
       (CheckChar = #146) Or {ע}
       (CheckChar = #147) Or {פ}
       (CheckChar = #148) Or {ף}
       (CheckChar = #149) Or {צ}
       (CheckChar = #150) Or {ץ}
       (CheckChar = #151) Or {ק}
       (CheckChar = #152) Or {ר}
       (CheckChar = #153) Or {ש}
       (CheckChar = #154) Or {ת}
       (UpperCase(CheckChar) = 'A') Or
       (UpperCase(CheckChar) = 'B') Or
       (UpperCase(CheckChar) = 'C') Or
       (UpperCase(CheckChar) = 'D') Or
       (UpperCase(CheckChar) = 'E') Or
       (UpperCase(CheckChar) = 'F') Or
       (UpperCase(CheckChar) = 'G') Or
       (UpperCase(CheckChar) = 'H') Or
       (UpperCase(CheckChar) = 'I') Or
       (UpperCase(CheckChar) = 'J') Or
       (UpperCase(CheckChar) = 'K') Or
       (UpperCase(CheckChar) = 'L') Or
       (UpperCase(CheckChar) = 'M') Or
       (UpperCase(CheckChar) = 'N') Or
       (UpperCase(CheckChar) = 'O') Or
       (UpperCase(CheckChar) = 'P') Or
       (UpperCase(CheckChar) = 'Q') Or
       (UpperCase(CheckChar) = 'R') Or
       (UpperCase(CheckChar) = 'S') Or
       (UpperCase(CheckChar) = 'T') Or
       (UpperCase(CheckChar) = 'U') Or
       (UpperCase(CheckChar) = 'V') Or
       (UpperCase(CheckChar) = 'W') Or
       (UpperCase(CheckChar) = 'X') Or
       (UpperCase(CheckChar) = 'Y') Or
       (UpperCase(CheckChar) = 'Z') Or
       (CheckChar = 'א') Or
       (CheckChar = 'ב') Or
       (CheckChar = 'ג') Or
       (CheckChar = 'ד') Or
       (CheckChar = 'ה') Or
       (CheckChar = 'ו') Or
       (CheckChar = 'ז') Or
       (CheckChar = 'ח') Or
       (CheckChar = 'ט') Or
       (CheckChar = 'י') Or
       (CheckChar = 'כ') Or
       (CheckChar = 'ך') Or
       (CheckChar = 'ל') Or
       (CheckChar = 'מ') Or
       (CheckChar = 'ם') Or
       (CheckChar = 'נ') Or
       (CheckChar = 'ן') Or
       (CheckChar = 'ס') Or
       (CheckChar = 'ע') Or
       (CheckChar = 'פ') Or
       (CheckChar = 'ף') Or
       (CheckChar = 'צ') Or
       (CheckChar = 'ץ') Or
       (CheckChar = 'ק') Or
       (CheckChar = 'ר') Or
       (CheckChar = 'ש') Or
       (CheckChar = 'ת') Or
       (CheckChar = '_') Or
       (CheckChar = '-') Or
       (CheckChar = '.') Or
       (CheckChar = '@') Or
       (CheckChar = '#') Or
       (CheckChar = '$') Or
       (CheckChar = '^') Or
       (CheckChar = '&') Or
       (CheckChar = '(') Or
       (CheckChar = ')') Or
       (CheckChar = '[') Or
       (CheckChar = ']') Or
       (CheckChar = '>') Or
       (CheckChar = '<') Or
       (CheckChar = ':') Or
       (CheckChar = '"') Or
       (CheckChar = ' ') then
    begin
      Result := Result + CheckChar;
    end;
  end;

  Result := Trim(Result);
end;

function  CleanDirFileName(Dir : String) : String;
Var
  CurChar : Integer;
  CheckStr : String;
  CheckChar : Char;


(*

Naming Conventions
The following fundamental rules enable applications to create and process valid names for files and directories, regardless of the file system:

Use a period to separate the base file name from the extension in the name of a directory or file.
Use a backslash (\) to separate the components of a path. The backslash divides the file name from the path to it, and one directory name from another directory name in a path. You cannot use a backslash in the name for the actual file or directory because it is a reserved character that separates the names into components.
Use a backslash as required as part of volume names, for example, the "C:\" in "C:\path\file" or the "\\server\share" in "\\server\share\path\file" for Universal Naming Convention (UNC) names. For more information about UNC names, see the Maximum Path Length Limitation section.
Do not assume case sensitivity. For example, consider the names OSCAR, Oscar, and oscar to be the same, even though some file systems (such as a POSIX-compliant file system) may consider them as different. Note that NTFS supports POSIX semantics for case sensitivity but this is not the default behavior. For more information, see CreateFile.
Volume designators (drive letters) are similarly case-insensitive. For example, "D:\" and "d:\" refer to the same volume.
Use any character in the current code page for a name, including Unicode characters and characters in the extended character set (128–255), except for the following:

The following reserved characters:

< (less than)
> (greater than)
: (colon)
" (double quote)
/ (forward slash)
\ (backslash)
| (vertical bar or pipe)
? (question mark)
* (asterisk)
Integer value zero, sometimes referred to as the ASCII NUL character.
Characters whose integer representations are in the range from 1 through 31, except for alternate data streams where these characters are allowed. For more information about file streams, see File Streams.
Any other character that the target file system does not allow.
Use a period as a directory component in a path to represent the current directory, for example ".\temp.txt". For more information, see Paths.
Use two consecutive periods (..) as a directory component in a path to represent the parent of the current directory, for example "..\temp.txt". For more information, see Paths.
Do not use the following reserved names for the name of a file:

CON, PRN, AUX, NUL, COM1, COM2, COM3, COM4, COM5, COM6, COM7, COM8, COM9, LPT1, LPT2, LPT3, LPT4, LPT5, LPT6, LPT7, LPT8, and LPT9. Also avoid these names followed immediately by an extension; for example, NUL.txt is not recommended. For more information, see Namespaces.

Do not end a file or directory name with a space or a period. Although the underlying file system may support such names, the Windows shell and user interface does not. However, it is acceptable to specify a period as the first character of a name. For example, ".temp".
Short vs. Long Names
A long file name is considered to be any file name that exceeds the short MS-DOS (also called 8.3) style naming convention. When you create a long file name, Windows may also create a short 8.3 form of the name, called the 8.3 alias or short name, and store it on disk also. This 8.3 aliasing can be disabled for performance reasons either systemwide or for a specified volume, depending on the particular file system.

Windows Server 2008, Windows Vista, Windows Server 2003 and Windows XP:  8.3 aliasing cannot be disabled for specified volumes until Windows 7 and Windows Server 2008 R2.

On many file systems, a file name will contain a tilde (~) within each component of the name that is too long to comply with 8.3 naming rules.

*)

begin
  Result := 'NoName';
  CheckStr := Trim(Dir);
  if CheckStr = '' then
    Exit;
  Result := '';
  for CurChar := 1 to Length(CheckStr) do
  begin
    CheckChar := CheckStr[CurChar];
//    MyAsc := Ord(CheckChar);
//    IF MyAsc > 600 Then
//      Beep;

    IF (CheckChar = #224) Then {א}
       CheckChar := 'א';
    IF (CheckChar = #225) Then {ב}
       CheckChar := 'ב';
    IF (CheckChar = #226) Then {ג}
       CheckChar := 'ג';
    IF (CheckChar = #227) Then {ד}
       CheckChar := 'ד';
    IF (CheckChar = #228) Then {ה}
       CheckChar := 'ה';
    IF (CheckChar = #229) Then {ו}
       CheckChar := 'ו';
    IF (CheckChar = #230) Then {ז}
       CheckChar := 'ז';
    IF (CheckChar = #231) Then {ח}
       CheckChar := 'ח';
    IF (CheckChar = #232) Then {ט}
       CheckChar := 'ט';
    IF (CheckChar = #233) Then {י}
       CheckChar := 'י';
    IF (CheckChar = #234) Then {כ}
       CheckChar := 'ך';
    IF (CheckChar = #235) Then {ך}
       CheckChar := 'כ';
    IF (CheckChar = #236) Then {ל}
       CheckChar := 'ל';
    IF (CheckChar = #237) Then {מ}
       CheckChar := 'ם';
    IF (CheckChar = #238) Then {ם}
       CheckChar := 'מ';
    IF (CheckChar = #239) Then {נ}
       CheckChar := 'ן';
    IF (CheckChar = #240) Then {ן}
       CheckChar := 'נ';
    IF (CheckChar = #241) Then {ס}
       CheckChar := 'ס';
    IF (CheckChar = #242) Then {ע}
       CheckChar := 'ע';
    IF (CheckChar = #243) Then {פ}
       CheckChar := 'ף';
    IF (CheckChar = #244) Then {ף}
       CheckChar := 'פ';
    IF (CheckChar = #245) Then {צ}
       CheckChar := 'ץ';
    IF (CheckChar = #246) Then {ץ}
       CheckChar := 'צ';
    IF (CheckChar = #247) Then {ק}
       CheckChar := 'ק';
    IF (CheckChar = #248) Then {ר}
       CheckChar := 'ר';
    IF (CheckChar = #249) Then {ש}
       CheckChar := 'ש';
    IF (CheckChar = #250) Then {ת}
       CheckChar := 'ת';

    if (CheckChar = '0') Or
       (CheckChar = '1') Or
       (CheckChar = '2') Or
       (CheckChar = '3') Or
       (CheckChar = '4') Or
       (CheckChar = '5') Or
       (CheckChar = '6') Or
       (CheckChar = '7') Or
       (CheckChar = '8') Or
       (CheckChar = '9') Or
       (CheckChar = #128) Or {א}
       (CheckChar = #129) Or {ב}
       (CheckChar = #130) Or {ג}
       (CheckChar = #131) Or {ד}
       (CheckChar = #132) Or {ה}
       (CheckChar = #133) Or {ו}
       (CheckChar = #134) Or {ז}
       (CheckChar = #135) Or {ח}
       (CheckChar = #136) Or {ט}
       (CheckChar = #137) Or {י}
       (CheckChar = #138) Or {כ}
       (CheckChar = #139) Or {ך}
       (CheckChar = #140) Or {ל}
       (CheckChar = #141) Or {מ}
       (CheckChar = #142) Or {ם}
       (CheckChar = #143) Or {נ}
       (CheckChar = #144) Or {ן}
       (CheckChar = #145) Or {ס}
       (CheckChar = #146) Or {ע}
       (CheckChar = #147) Or {פ}
       (CheckChar = #148) Or {ף}
       (CheckChar = #149) Or {צ}
       (CheckChar = #150) Or {ץ}
       (CheckChar = #151) Or {ק}
       (CheckChar = #152) Or {ר}
       (CheckChar = #153) Or {ש}
       (CheckChar = #154) Or {ת}
       (UpperCase(CheckChar) = UpperCase('A')) Or
       (UpperCase(CheckChar) = UpperCase('B')) Or
       (UpperCase(CheckChar) = UpperCase('C')) Or
       (UpperCase(CheckChar) = UpperCase('D')) Or
       (UpperCase(CheckChar) = UpperCase('E')) Or
       (UpperCase(CheckChar) = UpperCase('F')) Or
       (UpperCase(CheckChar) = UpperCase('G')) Or
       (UpperCase(CheckChar) = UpperCase('H')) Or
       (UpperCase(CheckChar) = UpperCase('I')) Or
       (UpperCase(CheckChar) = UpperCase('J')) Or
       (UpperCase(CheckChar) = UpperCase('K')) Or
       (UpperCase(CheckChar) = UpperCase('L')) Or
       (UpperCase(CheckChar) = UpperCase('M')) Or
       (UpperCase(CheckChar) = UpperCase('N')) Or
       (UpperCase(CheckChar) = UpperCase('O')) Or
       (UpperCase(CheckChar) = UpperCase('P')) Or
       (UpperCase(CheckChar) = UpperCase('Q')) Or
       (UpperCase(CheckChar) = UpperCase('R')) Or
       (UpperCase(CheckChar) = UpperCase('S')) Or
       (UpperCase(CheckChar) = UpperCase('T')) Or
       (UpperCase(CheckChar) = UpperCase('U')) Or
       (UpperCase(CheckChar) = UpperCase('V')) Or
       (UpperCase(CheckChar) = UpperCase('W')) Or
       (UpperCase(CheckChar) = UpperCase('X')) Or
       (UpperCase(CheckChar) = UpperCase('Y')) Or
       (UpperCase(CheckChar) = UpperCase('Z')) Or
       (CheckChar = 'א') Or
       (CheckChar = 'ב') Or
       (CheckChar = 'ג') Or
       (CheckChar = 'ד') Or
       (CheckChar = 'ה') Or
       (CheckChar = 'ו') Or
       (CheckChar = 'ז') Or
       (CheckChar = 'ח') Or
       (CheckChar = 'ט') Or
       (CheckChar = 'י') Or
       (CheckChar = 'כ') Or
       (CheckChar = 'ך') Or
       (CheckChar = 'ל') Or
       (CheckChar = 'מ') Or
       (CheckChar = 'ם') Or
       (CheckChar = 'נ') Or
       (CheckChar = 'ן') Or
       (CheckChar = 'ס') Or
       (CheckChar = 'ע') Or
       (CheckChar = 'פ') Or
       (CheckChar = 'ף') Or
       (CheckChar = 'צ') Or
       (CheckChar = 'ץ') Or
       (CheckChar = 'ק') Or
       (CheckChar = 'ר') Or
       (CheckChar = 'ש') Or
       (CheckChar = 'ת') Or
       (CheckChar = '_') Or
       (CheckChar = '-') Or
       (CheckChar = ' ') then
    begin
      Result := Result + CheckChar;
    end;
  end;

  if Trim(Result) = '' then
    Result := 'NoName';
end;

function  CleanUnPrintableChar(CheckString :String) : String;
Var
  CurChar : Integer;
  CheckStr : String;
  CheckChar : Char;
begin
  Result := 'NoName';
  CheckStr := Trim(CheckString);
  if CheckStr = '' then
    Exit;
  Result := '';
  for CurChar := 1 to Length(CheckStr) do
  begin
    CheckChar := CheckStr[CurChar];

    IF (CheckChar = #224) Then {א}
       CheckChar := 'א';
    IF (CheckChar = #225) Then {ב}
       CheckChar := 'ב';
    IF (CheckChar = #226) Then {ג}
       CheckChar := 'ג';
    IF (CheckChar = #227) Then {ד}
       CheckChar := 'ד';
    IF (CheckChar = #228) Then {ה}
       CheckChar := 'ה';
    IF (CheckChar = #229) Then {ו}
       CheckChar := 'ו';
    IF (CheckChar = #230) Then {ז}
       CheckChar := 'ז';
    IF (CheckChar = #231) Then {ח}
       CheckChar := 'ח';
    IF (CheckChar = #232) Then {ט}
       CheckChar := 'ט';
    IF (CheckChar = #233) Then {י}
       CheckChar := 'י';
    IF (CheckChar = #234) Then {כ}
       CheckChar := 'ך';
    IF (CheckChar = #235) Then {ך}
       CheckChar := 'כ';
    IF (CheckChar = #236) Then {ל}
       CheckChar := 'ל';
    IF (CheckChar = #237) Then {מ}
       CheckChar := 'ם';
    IF (CheckChar = #238) Then {ם}
       CheckChar := 'מ';
    IF (CheckChar = #239) Then {נ}
       CheckChar := 'ן';
    IF (CheckChar = #240) Then {ן}
       CheckChar := 'נ';
    IF (CheckChar = #241) Then {ס}
       CheckChar := 'ס';
    IF (CheckChar = #242) Then {ע}
       CheckChar := 'ע';
    IF (CheckChar = #243) Then {פ}
       CheckChar := 'ף';
    IF (CheckChar = #244) Then {ף}
       CheckChar := 'פ';
    IF (CheckChar = #245) Then {צ}
       CheckChar := 'ץ';
    IF (CheckChar = #246) Then {ץ}
       CheckChar := 'צ';
    IF (CheckChar = #247) Then {ק}
       CheckChar := 'ק';
    IF (CheckChar = #248) Then {ר}
       CheckChar := 'ר';
    IF (CheckChar = #249) Then {ש}
       CheckChar := 'ש';
    IF (CheckChar = #250) Then {ת}
       CheckChar := 'ת';

    if (CheckChar = '0') Or
       (CheckChar = '1') Or
       (CheckChar = '2') Or
       (CheckChar = '3') Or
       (CheckChar = '4') Or
       (CheckChar = '5') Or
       (CheckChar = '6') Or
       (CheckChar = '7') Or
       (CheckChar = '8') Or
       (CheckChar = '9') Or
       (CheckChar = #128) Or {א}
       (CheckChar = #129) Or {ב}
       (CheckChar = #130) Or {ג}
       (CheckChar = #131) Or {ד}
       (CheckChar = #132) Or {ה}
       (CheckChar = #133) Or {ו}
       (CheckChar = #134) Or {ז}
       (CheckChar = #135) Or {ח}
       (CheckChar = #136) Or {ט}
       (CheckChar = #137) Or {י}
       (CheckChar = #138) Or {כ}
       (CheckChar = #139) Or {ך}
       (CheckChar = #140) Or {ל}
       (CheckChar = #141) Or {מ}
       (CheckChar = #142) Or {ם}
       (CheckChar = #143) Or {נ}
       (CheckChar = #144) Or {ן}
       (CheckChar = #145) Or {ס}
       (CheckChar = #146) Or {ע}
       (CheckChar = #147) Or {פ}
       (CheckChar = #148) Or {ף}
       (CheckChar = #149) Or {צ}
       (CheckChar = #150) Or {ץ}
       (CheckChar = #151) Or {ק}
       (CheckChar = #152) Or {ר}
       (CheckChar = #153) Or {ש}
       (CheckChar = #154) Or {ת}
       (UpperCase(CheckChar) = UpperCase('A')) Or
       (UpperCase(CheckChar) = UpperCase('B')) Or
       (UpperCase(CheckChar) = UpperCase('C')) Or
       (UpperCase(CheckChar) = UpperCase('D')) Or
       (UpperCase(CheckChar) = UpperCase('E')) Or
       (UpperCase(CheckChar) = UpperCase('F')) Or
       (UpperCase(CheckChar) = UpperCase('G')) Or
       (UpperCase(CheckChar) = UpperCase('H')) Or
       (UpperCase(CheckChar) = UpperCase('I')) Or
       (UpperCase(CheckChar) = UpperCase('J')) Or
       (UpperCase(CheckChar) = UpperCase('K')) Or
       (UpperCase(CheckChar) = UpperCase('L')) Or
       (UpperCase(CheckChar) = UpperCase('M')) Or
       (UpperCase(CheckChar) = UpperCase('N')) Or
       (UpperCase(CheckChar) = UpperCase('O')) Or
       (UpperCase(CheckChar) = UpperCase('P')) Or
       (UpperCase(CheckChar) = UpperCase('Q')) Or
       (UpperCase(CheckChar) = UpperCase('R')) Or
       (UpperCase(CheckChar) = UpperCase('S')) Or
       (UpperCase(CheckChar) = UpperCase('T')) Or
       (UpperCase(CheckChar) = UpperCase('U')) Or
       (UpperCase(CheckChar) = UpperCase('V')) Or
       (UpperCase(CheckChar) = UpperCase('W')) Or
       (UpperCase(CheckChar) = UpperCase('X')) Or
       (UpperCase(CheckChar) = UpperCase('Y')) Or
       (UpperCase(CheckChar) = UpperCase('Z')) Or
       (CheckChar = 'א') Or
       (CheckChar = 'ב') Or
       (CheckChar = 'ג') Or
       (CheckChar = 'ד') Or
       (CheckChar = 'ה') Or
       (CheckChar = 'ו') Or
       (CheckChar = 'ז') Or
       (CheckChar = 'ח') Or
       (CheckChar = 'ט') Or
       (CheckChar = 'י') Or
       (CheckChar = 'כ') Or
       (CheckChar = 'ך') Or
       (CheckChar = 'ל') Or
       (CheckChar = 'מ') Or
       (CheckChar = 'ם') Or
       (CheckChar = 'נ') Or
       (CheckChar = 'ן') Or
       (CheckChar = 'ס') Or
       (CheckChar = 'ע') Or
       (CheckChar = 'פ') Or
       (CheckChar = 'ף') Or
       (CheckChar = 'צ') Or
       (CheckChar = 'ץ') Or
       (CheckChar = 'ק') Or
       (CheckChar = 'ר') Or
       (CheckChar = 'ש') Or
       (CheckChar = 'ת') Or
       (CheckChar = '_') Or
       (CheckChar = '-') Or
       (CheckChar = '~') Or
       (CheckChar = '!') Or
       (CheckChar = '@') Or
       (CheckChar = '#') Or
       (CheckChar = '$') Or
       (CheckChar = '%') Or
       (CheckChar = '^') Or
       (CheckChar = '&') Or
       (CheckChar = '*') Or
       (CheckChar = '(') Or
       (CheckChar = ')') Or
       (CheckChar = '+') Or
       (CheckChar = '=') Or
       (CheckChar = '}') Or
       (CheckChar = '{') Or
       (CheckChar = '[') Or
       (CheckChar = ']') Or
       (CheckChar = '"') Or
       (CheckChar = '''') Or
       (CheckChar = '|') Or
       (CheckChar = '\') Or
       (CheckChar = '/') Or
       (CheckChar = '?') Or
       (CheckChar = '.') Or
       (CheckChar = ',') Or
       (CheckChar = ';') Or
       (CheckChar = ':') Or
       (CheckChar = '<') Or
       (CheckChar = '>') Or
       (CheckChar = ' ') then
    begin
      Result := Result + CheckChar;
    end;
  end;

  if Trim(Result) = '' then
    Result := 'NoName';
end;

function  GetKeyString : String;
begin
  Result := FormatDateTime('yyyymmdd' , Now) +
            IntToStr(GetTickCount64) +
            FormatDateTime('hhnnsszzz' , Now);
end;

function CheckIsraelIdNumber(Id : Integer) : Boolean;
Var
  Str   : String;
  CurNo : Integer;
  IntNo : Integer;
  Sum   : Integer;
begin
//
//  אלגוריתם: בדיקת תקינות של סיפרת הביקורת
//
//  רושמים את 9 הספרות - אם יש צורך, משלימים באפסים מובילים
//
//  0   1   2   3   4   5   6   7   8
//
//  מתחת לכל מספר - החל מצד ימין - רושמים 1 או 2 לפי הסדר
//
//  1   2   3   4   5   6   7   8   9
//  1   2   1   2   1   2   1   2   1
//
//
//  מכפילים כל קבוצה - לדוגמא
//
//  0   0   8   1   9   0   7   8   7
//  1   2   1   2   1   2   1   2   1
//  ---------------------------------
//  0   0   8   2   9   0   7   16  7
//
//  מחברים את הספרות בכל תוצאה
//
//  0   0   8   2   9   0   7   16  7
//
//  0   0   8   2   9   0   7    7  7
//
//  מחברים את סה"כ התוצאות
//  0  +0  +8  +2  +9   +0 +7   +7 +7 = 40
//
//  חלוקה בעשר ללא שארית - תקין
//
//  אחרת - שגוי
//

  Result := False;
  Str := IntToStr(ID);
  if Length(Str) > 9 then
    Exit;

  Str := ZeroIntToStr(9,id);
  if Length(Str) <> 9 then
    Exit;

  Sum := 0;
  For CurNo := 1 To 9 Do
  begin
    case CurNo of
       1,3,5,7,9 : begin
             Sum := Sum + StrToInt(Str[CurNo]);
           end;
       2,4,6,8 : begin
             IntNo := StrToInt(Str[CurNo]) * 2;
             if IntNo <= 9 then
             begin
               Sum := Sum + IntNo;
             end
             else
             begin
               Str    := IntToStr(IntNo);
               IntNo  := StrToInt(Str[1]) + StrToInt(Str[2]);
               Sum := Sum + IntNo;
             end;
           end;
    end;
  end;

  Str := IntToStr(Sum);
  if Str[Length(Str)] = '0' then
    Result := True;
end;

function CheckIsraelCompId(CompId : Integer) : Boolean;
Var
  Bikoret : String;
  Str   : String;
  ChkId : String;
  CurNo : Integer;
  IntNo : Integer;
  Sum   : Integer;
begin

// ספרת ביקורת
// 1. רשום את הספרות ללא סיפרת הביקורת
// 2. מתחת לכל סיפרה, החל מ*ימין* רשום 2121212 החל ב 2.
// 3. הכפל כל ספרה עם הספרה שמתחתה
// 4. בכל מקרה של תוצאה גדולה מ 9 , חבר את ספרות התוצאה למשל אם התוצאה 12 אז 2+1=3
// 5. חבר את כל התוצאות.
// 6. את ספרת היחידות של התוצאה הזאת החסר מ 10 למשל אם התוצאה 42 אז 10-2 =8
// 7. קיבלתה את ספרת הביקורת.


  Result := False;
  Str := IntToStr(CompID);
  if Length(Str) > 9 then
    Exit;

  Str := ZeroIntToStr(9,CompId);
  if Length(Str) <> 9 then
    Exit;

  Bikoret := Str[Length(Str)];

  Sum := 0;
  For CurNo := 8 DownTo 1 Do
  begin
    case CurNo of
       1,3,5,7 : begin
             Sum := Sum + StrToInt(Str[CurNo]);
           end;
       2,4,6,8 : begin
             IntNo := StrToInt(Str[CurNo]) * 2;
             if IntNo <= 9 then
             begin
               Sum := Sum + IntNo;
             end
             else
             begin
               Sum    := Sum + 1;
               IntNo  := IntNo - 10;
               Sum := Sum + IntNo;
             end;
           end;
    end;
  end;

  Str := IntToStr(Sum);
  ChkId := Str[Length(Str)];
  if ChkId = '0' then
  begin
    Result := (Bikoret = '0');
  end
  else
  begin
    IntNo := StrToIntDef(ChkId,99);
    ChkId := IntToStr(10 - IntNo);
    Result := (ChkId = Bikoret);
  end;
end;

function IsMatch(const Input, Pattern: string): boolean;
begin
  Result := TRegEx.IsMatch(Input, Pattern);
end;

function IsValidEmailRegEx(const EmailAddress: string): boolean;
const
  EMAIL_REGEX = '^((?>[a-zA-Z\d!#$%&''*+\-/=?^_`{|}~]+\x20*|"((?=[\x01-\x7f])'
             +'[^"\\]|\\[\x01-\x7f])*"\x20*)*(?<angle><))?((?!\.)'
             +'(?>\.?[a-zA-Z\d!#$%&''*+\-/=?^_`{|}~]+)+|"((?=[\x01-\x7f])'
             +'[^"\\]|\\[\x01-\x7f])*")@(((?!-)[a-zA-Z\d\-]+(?<!-)\.)+[a-zA-Z]'
             +'{2,}|\[(((?(?<!\[)\.)(25[0-5]|2[0-4]\d|[01]?\d?\d))'
             +'{4}|[a-zA-Z\d\-]*[a-zA-Z\d]:((?=[\x01-\x7f])[^\\\[\]]|\\'
             +'[\x01-\x7f])+)\])(?(angle)>)$';
begin
  Result := IsMatch(EmailAddress, EMAIL_REGEX);
end;

function IsPasswordValid(const s: AnsiString): Boolean;
const
  C_Upcase = 1;
  C_Locase = 2;
  C_Digit = 4;
  C_SpecSym = 8;
  C_All = C_Upcase or C_Locase or C_Digit or C_SpecSym;
var
  i, keys: integer;
begin

  if Length(s) < 8 then begin
    Result := False;
    Exit;
  end;

  keys := 0;
  for i := 1 to Length(s) do
    case s[i] of
      'A'..'Z': keys := keys or C_Upcase;
      'a'..'z': keys := keys or C_Locase;
      '0'..'9': keys := keys or C_Digit;
      '!','#','%','&','*','@','^','$',
          '(',')','[',']','{','}','-','_','+','=',
          '~','|','?','>','<','/','\': keys := keys or C_SpecSym;
    end;

  Result := keys = C_All;
end;

function IsValidPhoneNumber(const APhoneNumber: string): Boolean;
var
  S: string;
  I: Integer;
  HyphenCount: Integer;
  HyphenPos: Integer;
  DigitCount: Integer;
begin
  Result := False;
  S := APhoneNumber; // Working copy

  // 1. Rule: Must start with '0' (zero). This also ensures the string isn't empty.
  if (Length(S) = 0) or (S[1] <> '0') then
    Exit;

  HyphenCount := 0;
  HyphenPos := 0;
  DigitCount := 0;

  // 2. Loop over all characters to count digits and check non-numeric characters
  for I := 1 to Length(S) do
  begin
    if S[I] in ['0'..'9'] then
    begin
      Inc(DigitCount);
    end
    else // Found a non-numeric character
    begin
      // It MUST be a hyphen
      if S[I] = '-' then
      begin
        Inc(HyphenCount);
        // Rule: Only one hyphen allowed, check early
        if HyphenCount > 1 then
          Exit;

        HyphenPos := I;
      end
      else
      begin
        // Found a character that is NOT a digit AND NOT a hyphen
        Exit; // Invalid character, fail immediately
      end;
    end;
  end;

  // 3. Check validity based on hyphen presence and final digit count

  case HyphenCount of
    0: // No hyphen found (e.g., "0501234567" or "037615254")
      begin
        // Digits must be 9 (0n-nnnnnnn equivalent) or 10 (0nn-nnnnnnn equivalent)
        if (DigitCount = 9) or (DigitCount = 10) then
          Result := True;
      end;

    1: // Exactly one hyphen found
      begin
        // Rule: Hyphen must be in position 3 or 4
        if (HyphenPos = 3) then
        begin
          // Format: 0n-nnnnnnn (Total 10 characters, 9 digits)
          if DigitCount = 9 then
            Result := True;
        end
        else if (HyphenPos = 4) then
        begin
          // Format: 0nn-nnnnnnn (Total 11 characters, 10 digits)
          if DigitCount = 10 then
            Result := True;
        end;
      end;
  end;

end;

function  GetRegionByPrefix(Prefix : integer) : String;
begin
  Result := 'IL';
  case Prefix of
    1   : Result := 'US'; // "US","AG","AI","AS","BB","BM","BS","CA",
                          // "DM","DO","GD","GU","JM","KN","KY","LC",
                          // "MP","MS","PR","SX","TC","TT","VC","VG","VI"
    7   : Result := 'RU'; // "RU", "KZ"
    20  : Result := 'EG';
    27  : Result := 'ZA';
    30  : Result := 'GR';
    31  : Result := 'NL';
    32  : Result := 'BE';
    33  : Result := 'FR';
    34  : Result := 'ES';
    36  : Result := 'HU';
    39  : Result := 'IT'; //"IT", "VA',
    40  : Result := 'RO';
    41  : Result := 'CH';
    43  : Result := 'AT';
    44  : Result := 'GB'; // "GB", "GG", "IM", "JE',
    45  : Result := 'DK';
    46  : Result := 'SE';
    47  : Result := 'NO'; // "NO", "SJ',
    48  : Result := 'PL';
    49  : Result := 'DE';
    51  : Result := 'PE';
    52  : Result := 'MX';
    53  : Result := 'CU';
    54  : Result := 'AR';
    55  : Result := 'BR';
    56  : Result := 'CL';
    57  : Result := 'CO';
    58  : Result := 'VE';
    60  : Result := 'MY';
    61  : Result := 'AU'; //"AU", "CC", "CX',
    62  : Result := 'ID';
    63  : Result := 'PH';
    64  : Result := 'NZ';
    65  : Result := 'SG';
    66  : Result := 'TH';
    81  : Result := 'JP';
    82  : Result := 'KR';
    84  : Result := 'VN';
    86  : Result := 'CN';
    90  : Result := 'TR';
    91  : Result := 'IN';
    92  : Result := 'PK';
    93  : Result := 'AF';
    94  : Result := 'LK';
    95  : Result := 'MM';
    98  : Result := 'IR';
    211 : Result := 'SS';
    212 : Result := 'MA'; //"MA", "EH',
    213 : Result := 'DZ';
    216 : Result := 'TN';
    218 : Result := 'LY';
    220 : Result := 'GM';
    221 : Result := 'SN';
    222 : Result := 'MR';
    223 : Result := 'ML';
    224 : Result := 'GN';
    225 : Result := 'CI';
    226 : Result := 'BF';
    227 : Result := 'NE';
    228 : Result := 'TG';
    229 : Result := 'BJ';
    230 : Result := 'MU';
    231 : Result := 'LR';
    232 : Result := 'SL';
    233 : Result := 'GH';
    234 : Result := 'NG';
    235 : Result := 'TD';
    236 : Result := 'CF';
    237 : Result := 'CM';
    238 : Result := 'CV';
    239 : Result := 'ST';
    240 : Result := 'GQ';
    241 : Result := 'GA';
    242 : Result := 'CG';
    243 : Result := 'CD';
    244 : Result := 'AO';
    245 : Result := 'GW';
    246 : Result := 'IO';
    247 : Result := 'AC';
    248 : Result := 'SC';
    249 : Result := 'SD';
    250 : Result := 'RW';
    251 : Result := 'ET';
    252 : Result := 'SO';
    253 : Result := 'DJ';
    254 : Result := 'KE';
    255 : Result := 'TZ';
    256 : Result := 'UG';
    257 : Result := 'BI';
    258 : Result := 'MZ';
    260 : Result := 'ZM';
    261 : Result := 'MG';
    262 : Result := 'RE'; // "RE", "YT',
    263 : Result := 'ZW';
    264 : Result := 'NA';
    265 : Result := 'MW';
    266 : Result := 'LS';
    267 : Result := 'BW';
    268 : Result := 'SZ';
    269 : Result := 'KM';
    290 : Result := 'SH'; // "SH", "TA',
    291 : Result := 'ER';
    297 : Result := 'AW';
    298 : Result := 'FO';
    299 : Result := 'GL';
    350 : Result := 'GI';
    351 : Result := 'PT';
    352 : Result := 'LU';
    353 : Result := 'IE';
    354 : Result := 'IS';
    355 : Result := 'AL';
    356 : Result := 'MT';
    357 : Result := 'CY';
    358 : Result := 'FI'; // "FI", "AX',
    359 : Result := 'BG';
    370 : Result := 'LT';
    371 : Result := 'LV';
    372 : Result := 'EE';
    373 : Result := 'MD';
    374 : Result := 'AM';
    375 : Result := 'BY';
    376 : Result := 'AD';
    377 : Result := 'MC';
    378 : Result := 'SM';
    380 : Result := 'UA';
    381 : Result := 'RS';
    382 : Result := 'ME';
    383 : Result := 'XK';
    385 : Result := 'HR';
    386 : Result := 'SI';
    387 : Result := 'BA';
    389 : Result := 'MK';
    420 : Result := 'CZ';
    421 : Result := 'SK';
    423 : Result := 'LI';
    500 : Result := 'FK';
    501 : Result := 'BZ';
    502 : Result := 'GT';
    503 : Result := 'SV';
    504 : Result := 'HN';
    505 : Result := 'NI';
    506 : Result := 'CR';
    507 : Result := 'PA';
    508 : Result := 'PM';
    509 : Result := 'HT';
    590 : Result := 'GP'; //GP", "BL", "MF',
    591 : Result := 'BO';
    592 : Result := 'GY';
    593 : Result := 'EC';
    594 : Result := 'GF';
    595 : Result := 'PY';
    596 : Result := 'MQ';
    597 : Result := 'SR';
    598 : Result := 'UY';
    599 : Result := 'CW'; //"CW", "BQ',
    670 : Result := 'TL';
    672 : Result := 'NF';
    673 : Result := 'BN';
    674 : Result := 'NR';
    675 : Result := 'PG';
    676 : Result := 'TO';
    677 : Result := 'SB';
    678 : Result := 'VU';
    679 : Result := 'FJ';
    680 : Result := 'PW';
    681 : Result := 'WF';
    682 : Result := 'CK';
    683 : Result := 'NU';
    685 : Result := 'WS';
    686 : Result := 'KI';
    687 : Result := 'NC';
    688 : Result := 'TV';
    689 : Result := 'PF';
    690 : Result := 'TK';
    691 : Result := 'FM';
    692 : Result := 'MH';
    850 : Result := 'KP';
    852 : Result := 'HK';
    853 : Result := 'MO';
    855 : Result := 'KH';
    856 : Result := 'LA';
    880 : Result := 'BD';
    886 : Result := 'TW';
    960 : Result := 'MV';
    961 : Result := 'LB';
    962 : Result := 'JO';
    963 : Result := 'SY';
    964 : Result := 'IQ';
    965 : Result := 'KW';
    966 : Result := 'SA';
    967 : Result := 'YE';
    968 : Result := 'OM';
    970 : Result := 'PS';
    971 : Result := 'AE';
    972 : Result := 'IL';
    973 : Result := 'BH';
    974 : Result := 'QA';
    975 : Result := 'BT';
    976 : Result := 'MN';
    977 : Result := 'NP';
    992 : Result := 'TJ';
    993 : Result := 'TM';
    994 : Result := 'AZ';
    995 : Result := 'GE';
    996 : Result := 'KG';
    998 : Result := 'UZ';
  end;
end;

function  CutStringToWords(FullString : String; WordList : TStringList; DoSort : Boolean = True) : Boolean;
Var
  StrDynArr : TStringDynArray;
  CurItem : Integer;
  Str : String;
begin
  Result := False;
  WordList.Clear;
  FullString := RemoveAllEnterChar(FullString);
  StrDynArr := SplitString(FullString,' ');
  For CurItem := Low(StrDynArr) to High(StrDynArr) do
  begin
    Str := Trim(StrDynArr[CurItem]);
    if Str <> '' then
      WordList.Add(Str);
  end;

  if DoSort then
    WordList.Sort;
  Result := (WordList.Count > 0);
end;

function  CutStringToWords(FullString : String; WordList : TStringList; Delimiter : Char; DoSort : Boolean = True) : Boolean;
Var
  StrDynArr : TStringDynArray;
  CurItem : Integer;
  Str : String;
begin
  Result := False;
  WordList.Clear;
  FullString := RemoveAllEnterChar(FullString);
  StrDynArr := SplitString(FullString,String(Delimiter));
  For CurItem := Low(StrDynArr) to High(StrDynArr) do
  begin
    Str := Trim(StrDynArr[CurItem]);
    if Str <> '' then
      WordList.Add(Str);
  end;

  if DoSort then
    WordList.Sort;
  Result := (WordList.Count > 0);
end;

function  RemoveDuplicateWords(SearchText : String): String;
Var
  CurWrd        : Integer;
  WordList      : TStringList;
  NewWordList   : TStringList;
  StrTmp        : String;
  NewSearchText : String;
begin
  Result := SearchText;
  if Trim(SearchText) = '' then
    Exit;

  NewSearchText := SearchText;
  NewWordList := TStringList.Create;
  WordList    := TStringList.Create;
  Try
    IF CutStringToWords(SearchText,WordList,False) Then
    begin
      if WordList.Count > 1 then
      begin
        NewSearchText := '';
        For CurWrd := 0 To WordList.Count - 1 Do
        begin
          StrTmp := Trim(WordList.Strings[CurWrd]);
          if NewWordList.IndexOf(StrTmp) < 0 then
          begin
            NewWordList.Add(StrTmp);
          end;
        end;

        For CurWrd := 0 To NewWordList.Count - 1 Do
        begin
          NewSearchText := NewSearchText + ' ' + NewWordList.Strings[CurWrd];
        end;
        Result := NewSearchText;
      end;
    end;
  Finally
    WordList.Free;
    NewWordList.Free;
  End;
end;

function  AddMsgToEventLog(FMachine     : String;
                           FEventSource : String;
                           EventMsg     : String;
                           EventType    : Integer;
                           Category     : Integer;
                           MessageID    : Integer) : Boolean;
Var
  fLogHandle   : THandle;
  MessageStr   : array of PChar;
  MessageCount : Word;
begin

(*
Event Type

 1 EVENTLOG_ERROR_TYPE	        Error event
 2 EVENTLOG_WARNING_TYPE	Warning event
 4 EVENTLOG_INFORMATION_TYPE	Information event
 8 EVENTLOG_AUDIT_SUCCESS	Success Audit event
16 EVENTLOG_AUDIT_FAILURE	Failure Audit event

Category (Word):
    $0 = None
    $1 = Devices
    $2 = Disk
    $3 = Printers
    $4 = Services
    $5 = Shell
    $6 = System Event
    $7 = Network

  EVENTLOG_CATEGORY_NONE         : word = $0000;
  EVENTLOG_CATEGORY_DEVICES      : word = $0001;
  EVENTLOG_CATEGORY_DISK         : word = $0002;
  EVENTLOG_CATEGORY_PRINTERS     : word = $0003;
  EVENTLOG_CATEGORY_SERVICES     : word = $0004;
  EVENTLOG_CATEGORY_SHELL        : word = $0005;
  EVENTLOG_CATEGORY_SYSTEM_EVENT : word = $0006;
  EVENTLOG_CATEGORY_NETWORK      : word = $0007;

*)

  (*
  MyMessage     := TStringList.Create;
  Try
    MyMessage.Add('Line 01');
    aMessageCount := MyMessage.Count;
    SetLength(aMessageStr, aMessageCount);
    for I := 0 to aMessageCount - 1 do
        aMessageStr[I] := StrNew(PChar(MyMessage.Strings[I]));

  Finally
    MyMessage.Free;
  End;
  *)

  Result := True;

  MessageCount := 1;
  SetLength(MessageStr, MessageCount);
  MessageStr[0] := Pchar(EventMsg);

  Try
    IF FMachine = '' Then
      fLogHandle := Windows.RegisterEventSource(nil, PChar(FEventSource))
    else
      fLogHandle := Windows.RegisterEventSource(PChar(FMachine), PChar(FEventSource));

    Try
      Windows.ReportEvent(fLogHandle,      // This handle Identifies the event log.
                          EventType,       // Specifies the type of event being logged
                          Category,        // Specifies the event category.
                          MessageID,       // Specifies the event identifier
                          nil,             // Points to the current user's security identifier
                          MessageCount,    // Specifies the number of strings in the Message array
                          0,               // Specifies the number of bytes of event-specific raw (binary) data (like into %1..)
                          MessageStr,      // Points to a buffer containing an array of null-terminated strings that are merged into the message
                          nil);            // Points to the buffer containing the binary data

      //Windows.ReportEvent( fLogHandle, aEventType, aCategoryID, aEventID, nil, aMessageCount,
      //        Length(aData), aMessageStr, @aData[1]);

      //SvcMgr.LogMessage('This is just string', aEventType, aCategoryID, aEventID)

    Finally
      DeregisterEventSource(fLogHandle);
    End;
  Except;
    Result := False;
  End;
end;

function ConvertFileToBase64_(FileName : String; var Base64Str : String) : Boolean;
var
  fData  : HCkBinData;
  success  : Boolean;
  b64      : PWideChar;
begin
  Result := False;
  Base64Str := '';

  Try
    fData := CkBinData_Create();

    success := CkBinData_LoadFile(fData,PWideChar(FileName));
    if (success <> True) then
    begin
      AddMsgToEventLog('', J_FirstFileName(ParamStr(0)), 'Fail Load Pdf File [' + FileName + '] to Base64',EVENTLOG_ERROR_TYPE, 4, 1);
      Result := False;
      Exit;
    end;

    // Encode the PDF to base64
    // Note: to produce base64 on multiple lines (as it would appear in the MIME of an email),
    // pass the string "base64_mime" instead of "base64".
    b64 := CkBinData__getEncoded(fData,'base64');
    Base64Str := String(b64);

    Result := True;
  Finally
    CkBinData_Dispose(fData);
    fData := nil;
  End;
end;

function ConvertFileToBase64Mime_(FileName : String; var Base64StrMime : String) : Boolean;
var
  fData  : HCkBinData;
  success  : Boolean;
  b64      : PWideChar;
begin
  Result := False;
  Base64StrMime := '';

  Try
    fData := CkBinData_Create();

    success := CkBinData_LoadFile(fData,PWideChar(FileName));
    if (success <> True) then
    begin
      AddMsgToEventLog('', J_FirstFileName(ParamStr(0)), 'Fail Load File [' + FileName + '] to Base64Mime',EVENTLOG_ERROR_TYPE, 4, 1);
      Result := False;
      Exit;
    end;

    // Encode the PDF to base64
    // Note: to produce base64 on multiple lines (as it would appear in the MIME of an email),
    // pass the string "base64_mime" instead of "base64".
    b64 := CkBinData__getEncoded(fData,'base64_mime');
    Base64StrMime := String(b64);

    Result := True;
  Finally
    CkBinData_Dispose(fData);
    fData := nil;
  End;
end;

function ConvertBase64ToFile_(Base64Str : String; FileName : String) : Boolean;
var
  fData  : HCkBinData;
  success  : Boolean;
begin
  Result := False;

  Try
    fData := CkBinData_Create();

    // Decode from base64 PDF.
    fData := CkBinData_Create();
    CkBinData_AppendEncoded(fData,PWideChar(Base64Str),'base64');
    success := CkBinData_WriteFile(fData,PWideChar(FileName));
    if (success <> True) then
    begin
      AddMsgToEventLog('', J_FirstFileName(ParamStr(0)), 'Fail Decode Pdf File [' + FileName + '] From Base64',EVENTLOG_ERROR_TYPE, 4, 1);
      Exit;
    end;

    Result := True;
  Finally
    CkBinData_Dispose(fData);
    fData := nil;
  End;
end;

function ConvertBase64ToFileMime_(Base64StrMime : String; FileName : String) : Boolean;
var
  fData  : HCkBinData;
  success  : Boolean;
begin
  Result := False;

  Try
    fData := CkBinData_Create();

    // Decode from base64 PDF.
    fData := CkBinData_Create();
    CkBinData_AppendEncoded(fData,PWideChar(Base64StrMime),'base64_mime');
    success := CkBinData_WriteFile(fData,PWideChar(FileName));
    if (success <> True) then
    begin
      AddMsgToEventLog('', J_FirstFileName(ParamStr(0)), 'Fail Decode Pdf File [' + FileName + '] From Base64Mime',EVENTLOG_ERROR_TYPE, 4, 1);
      Exit;
    end;

    Result := True;
  Finally
    CkBinData_Dispose(fData);
    fData := nil;
  End;
end;

function FileToBase64_(const FilePath: string): string;
var
  FileStream: TMemoryStream;
  Bytes: TBytes;
begin
  Result := '';
  if not FileExists(FilePath) then
    raise Exception.Create('File not found: ' + FilePath);

  FileStream := TMemoryStream.Create;
  try
    // Load the file into the memory stream
    FileStream.LoadFromFile(FilePath);

    // Read the contents of the memory stream into a byte array
    SetLength(Bytes, FileStream.Size);
    FileStream.Position := 0;
    FileStream.Read(Bytes[0], FileStream.Size);

    // Convert to Base64
    Result := TNetEncoding.Base64.EncodeBytesToString(Bytes);
  finally
    FileStream.Free;
  end;
end;

function Base64ToFile_(const Base64String, FilePath: string) : Boolean;
var
  Base64Str  : String;
  FileStream: TFileStream;
  Bytes: TBytes;
  ImageStartPos : Integer;

begin
  Result := False;
  Base64Str  := Base64String;
  // Remove the prefix "data:image/png;base64," from the Base64 string
  ImageStartPos := Pos(',', Base64Str);
  if ImageStartPos > 0 then
    Base64Str := Copy(Base64Str, ImageStartPos + 1, MaxInt);

  // Decode Base64 string to bytes
  Bytes := TNetEncoding.Base64.DecodeStringToBytes(Base64Str);

  Try
    // Write bytes to file
    FileStream := TFileStream.Create(FilePath, fmCreate);
    try
      FileStream.WriteBuffer(Bytes[0], Length(Bytes));
      Result := True;
    finally
      FileStream.Free;
    end;
  Except;
    Result := False;
  End;
end;

function TColorToHTMLColor(Color: TColor): string;
var
  RGBColor: Longint;
  R, G, B: Byte;
begin
  // Convert the TColor value to an RGB color
  RGBColor := ColorToRGB(Color);

  // Extract the Red, Green, and Blue components
  R := GetRValue(RGBColor);
  G := GetGValue(RGBColor);
  B := GetBValue(RGBColor);

  // Format as a hex string (#RRGGBB)
  Result := Format('#%.2x%.2x%.2x', [R, G, B]);
end;

function IsSQLCommandDanger(SQLCOMMAND : String) : Boolean;
Var
  SqlStr    : String;
  fPassWord : String;
  SimpleSelect : Boolean;
Const
  SqlCheckStr01 : String = 'SELECT';
  SqlCheckStr02 : String = 'UPDATE';
  SqlCheckStr03 : String = 'DROP';
  SqlCheckStr04 : String = 'DELETE';
  SqlCheckStr05 : String = 'TRUNCATE';
  SqlCheckStr06 : String = 'ALTER';
  SqlCheckStr07 : String = 'CREATE';
begin
  Result := True;
  SqlStr := Trim(SQLCOMMAND);
  IF SqlStr = '' Then
    Exit;

  SimpleSelect := True;
  If UpperCase(Copy(SqlStr,1,Length(SqlCheckStr01))) <> UpperCase(SqlCheckStr01) Then
  begin
    // look for first characters 'SELECT'
    SimpleSelect := False;
  end;

  If TRegEx.IsMatch(UpperCase(SqlStr), '\b(DELETE|UPDATE|TRUNCATE|DROP|ALTER|CREATE)\b') Then
  begin
    SimpleSelect := False;
  end;

  if SimpleSelect Then
  begin
    Result := False;
    Exit;
  end;

  Result := True;
end;


end.
