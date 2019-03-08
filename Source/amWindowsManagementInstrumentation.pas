unit amWindowsManagementInstrumentation;
//******************************************************************************
//
//      Windows Management Interface
//
//******************************************************************************

(*

Examples of usage:

// Call a method on an instance and extract the result.
var
  WindowsManagement: IWindowsManagement;
  Method: IWindowsManagementObjectMethod;
  MethodResult: IWindowsManagementObjectMethodResult;
  Script: string;
const
  Scope = '\root\Microsoft\SqlServer\ReportServer\RS_SSRS\V14\Admin';
  ClassName = 'MSReportServer_ConfigurationSetting';
  MethodName = 'GenerateDatabaseCreationScript';
begin
  WindowsManagement := ConnectWindowsManagement(Scope);
  try
    Method := WindowsManagement.Objects[ClassName].Instances.First.Methods[MethodName];

    Method.Params['DatabaseName'] := 'ReportServer';
    Method.Params['Lcid'] := 1033;
    Method.Params['IsSharePointMode'] := False;

    MethodResult := Method.Execute;

    Script := VarToStr(MethodResult['Script']);

    ShowMessage(Script);
  finally
    WindowsManagement := nil;
  end;
end;

// Same as above using anonymous method parameters.
var
  WindowsManagement: IWindowsManagement;
  Instance: IWindowsManagementInstanceObject;
  Method: IWindowsManagementObjectMethod;
  MethodResult: IWindowsManagementObjectMethodResult;
  Script: string;
const
  Scope = '\root\Microsoft\SqlServer\ReportServer\RS_SSRS\V14\Admin';
  ClassName = 'MSReportServer_ConfigurationSetting';
  MethodName = 'GenerateDatabaseCreationScript';
begin
  WindowsManagement := ConnectWindowsManagement(Scope);
  try
    Instance := WindowsManagement.Objects[ClassName].Instances.First;

    Script := VarToStr(Instance.Methods[MethodName].Execute(['ReportServer', 1033, False])['Script']);

    ShowMessage(Script);
  finally
    WindowsManagement := nil;
  end;
end;

// Same as above using late binding via custom variant type.
var
  WindowsManagement: IWindowsManagement;
  Instance: Variant;
  Script: string;
const
  Scope = '\root\Microsoft\SqlServer\ReportServer\RS_SSRS\V14\Admin';
  ClassName = 'MSReportServer_ConfigurationSetting';
begin
  WindowsManagement := ConnectWindowsManagement(Scope);
  try
    Instance := WindowsManagementObject(WindowsManagement.Objects[ClassName].Instances.First);

    Script := VarToStr(Instance.GenerateDatabaseCreationScript('ReportServer', 1033, False).Script);

    ShowMessage(Script);
  finally
    WindowsManagement := nil;
  end;
end;

// Same as above using late binding via IDispatch intercept.
var
  WindowsManagement: IWindowsManagement;
  Instance: Variant;
  Script: string;
const
  Scope = '\root\Microsoft\SqlServer\ReportServer\RS_SSRS\V14\Admin';
  ClassName = 'MSReportServer_ConfigurationSetting';
begin
  WindowsManagement := ConnectWindowsManagement(Scope);
  try
    Instance := WindowsManagement.Objects[ClassName].Instances.First;

    Script := VarToStr(Instance.GenerateDatabaseCreationScript('ReportServer', 1033, False).Script);

    ShowMessage(Script);
  finally
    WindowsManagement := nil;
  end;
end;


// Execute a WMI query and extract values from the result.
var
  WindowsManagement: IWindowsManagement;
  Instances: IWindowsManagementInstanceList;
  Instance: IWindowsManagementObject;
const
  Scope = 'root\cimv2';
  Query := 'SELECT * FROM Win32_OperatingSystem';
begin
  WindowsManagement := ConnectWindowsManagement(Scope);
  try
    Instances := WindowsManagement.ExecuteQuery(Query);

    for Instance in Instances do
      ShowMessage(Instance['Name']);

  finally
    WindowsManagement := nil;
  end;
end;


*)

interface

uses
  JwaWbemCli,
  ComObj,
  SysUtils,
  Windows;

//------------------------------------------------------------------------------
//
//      Method, parameter and result objects
//
//------------------------------------------------------------------------------
type
  IWindowsManagementObjectMethodResult = interface
    ['{6F84FC0A-E277-4641-A80C-31134443031E}']
    function GetParam(const AName: string): Variant;

    property Params[const AName: string]: Variant read GetParam; default;
  end;

  IWindowsManagementObjectMethod = interface
    ['{9455F3F9-5CB7-4877-8017-68BDCC662F62}']
    procedure SetParam(const AName: string; const Value: Variant);
    function GetIsStatic: boolean;

    property IsStatic: boolean read GetIsStatic;
    property Params[const AName: string]: Variant write SetParam; default;
    function Execute: IWindowsManagementObjectMethodResult; overload;
    function Execute(const Params: array of Variant): IWindowsManagementObjectMethodResult; overload;
  end;

//------------------------------------------------------------------------------
//
//      Class and Instance wrapper
//
//------------------------------------------------------------------------------
type
  IWindowsManagementClassObject = interface;

  // Base interface common to both class and instance
  IWindowsManagementObject = interface
    ['{0AFD6DC4-D44C-42B3-9429-DEC0BF7330B8}']
    function GetNativeObject: IWbemClassObject;
    function GetValue(const AName: string): Variant;
    procedure SetValue(const AName: string; const AValue: Variant);
    function GetMethod(const AMethodName: string): IWindowsManagementObjectMethod;
    function GetIsClass: boolean;
    function GetIsInstance: boolean;
    function GetIsSingleton: boolean;
    function GetClassName: string;
    function GetKeyName: string;
    function GetPath: string;
    function GetRelativePath: string;
    function GetClassObject: IWindowsManagementClassObject;

    // Provide access to the wrapped WMI class object interface pointer
    property NativeObject: IWbemClassObject read GetNativeObject;

    property IsClass: boolean read GetIsClass;
    property IsInstance: boolean read GetIsInstance;
    property IsSingleton: boolean read GetIsSingleton;

    property ClassObject: IWindowsManagementClassObject read GetClassObject;
    property ClassName: string read GetClassName;
    property KeyName: string read GetKeyName;
    property Path: string read GetPath;
    property RelativePath: string read GetRelativePath;

    property Values[const AName: string]: Variant read GetValue write SetValue; default;
    property Methods[const AMethodName: string]: IWindowsManagementObjectMethod read GetMethod;
  end;

//------------------------------------------------------------------------------
//
//      Classes and related lists and enumerators
//
//------------------------------------------------------------------------------
  IWindowsManagementInstanceList = interface;
  IWindowsManagementInstanceObject = interface;

  // A class
  IWindowsManagementClassObject = interface(IWindowsManagementObject)
    ['{30B2F88D-0148-4018-B597-78D416E528DC}']
    function GetInstances: IWindowsManagementInstanceList;
    function GetInstance: IWindowsManagementInstanceObject;

    property Instance: IWindowsManagementInstanceObject read GetInstance; // Singleton instance
    property Instances: IWindowsManagementInstanceList read GetInstances; // List of instances
  end;

  // A class enumerator
  IWindowsManagementClassEnumerator = interface
    ['{DA6A0CEA-CE2D-4BD7-B382-FCBE75B96A1D}']
    function GetCurrent: IWindowsManagementClassObject;
    property Current: IWindowsManagementClassObject read GetCurrent;
    function MoveNext: Boolean;
  end;

  // A list of classes with random access
  IWindowsManagementClassList = interface
    ['{C667A395-258C-4537-9700-5386BE74A162}']
    function GetEnumerator: IWindowsManagementClassEnumerator;

    function GetItem(const APath: string): IWindowsManagementClassObject;
    property Items[const APath: string]: IWindowsManagementClassObject read GetItem; default;
  end;


//------------------------------------------------------------------------------
//
//      Instances and related lists and enumerators
//
//------------------------------------------------------------------------------
  IWindowsManagementInstanceObject = interface(IWindowsManagementObject)
    ['{DE4705A0-FE1D-4849-8ED3-6BACD71FBFE6}']
    function GetInstanceName: string;
    property InstanceName: string read GetInstanceName;
  end;

  // An instance enumerator
  IWindowsManagementInstanceEnumerator = interface
    ['{156069B9-F507-4EAE-A17E-395D6A5B8F64}']
    function GetCurrent: IWindowsManagementInstanceObject;
    property Current: IWindowsManagementInstanceObject read GetCurrent;
    function MoveNext: Boolean;
  end;

  // A simple list of instances (just wraps the enumerator)
  IWindowsManagementSimpleInstanceList = interface
    ['{A2CEEE56-7C7A-4B19-A525-9319AC73E9AC}']
    function GetFirst: IWindowsManagementInstanceObject;

    property First: IWindowsManagementInstanceObject read GetFirst;
    function GetEnumerator: IWindowsManagementInstanceEnumerator;
  end;

  // A list of instances with random access
  IWindowsManagementInstanceList = interface(IWindowsManagementSimpleInstanceList)
    ['{93B15424-19AC-4AE3-803D-810DFF6A68BC}']
    function GetItem(const AInstanceName: string): IWindowsManagementInstanceObject;
    property Items[const AInstanceName: string]: IWindowsManagementInstanceObject read GetItem; default;
  end;


//------------------------------------------------------------------------------
//
//      The main WMI namespace/scope/connection object
//
//------------------------------------------------------------------------------
type
  IWindowsManagement = interface
    ['{C6F2C902-9293-4E83-AAEF-BC7D77CA74B5}']
    function GetObjects: IWindowsManagementClassList;

    function ExecuteQuery(const Query: string): IWindowsManagementSimpleInstanceList;

    property Objects: IWindowsManagementClassList read GetObjects;
  end;


//------------------------------------------------------------------------------
//
//      Main entry point/factory
//
//------------------------------------------------------------------------------
function ConnectWindowsManagement(const AScope: string = ''; const AComputerName: string = 'localhost'; const AUserName: string = ''; const APassword: string = ''; const AContext: IWbemContext = nil): IWindowsManagement;

procedure WindowsManagementInitializeSecurity(LocalConnection: boolean);


//------------------------------------------------------------------------------
//
//      Error handling
//
//------------------------------------------------------------------------------
// WmiCheck() is used to check the result of a WMI API call.
// It throws an EWMIError exception on error.
//------------------------------------------------------------------------------
procedure WmiCheck(ErrorCode: HRESULT);

type
  EWMIError = class(EOleSysError)
  public
    constructor Create(ErrorCode: HRESULT);
  end;


//------------------------------------------------------------------------------
//
//      Custom variants
//
//------------------------------------------------------------------------------
// Wrap an IWindowsManagementObject in an invokable custom variant.
//------------------------------------------------------------------------------
function WindowsManagementObject(const AWindowsManagementObject: IWindowsManagementObject = nil): Variant;


//------------------------------------------------------------------------------
//
//      Variant array enumerator
//
//------------------------------------------------------------------------------
// Usage:
// var
//   VarArray: Variant;
//   Value: Variant;
// begin
//   for Value in OleVariantArrayEnum(VarArray) do
//     ...do something with Value...
// end;
//------------------------------------------------------------------------------
type
  IOleVariantEnum  = interface
    function GetCurrent: OLEVariant;
    function MoveNext: Boolean;
    property Current: OLEVariant read GetCurrent;
  end;

  IGetOleVariantEnum = interface
    function GetEnumerator: IOleVariantEnum;
  end;

// Make an IEnumVARIANT enumerable
function OleVariantEnum(const Collection: OleVariant): IGetOleVariantEnum;

// Make a variant array enumerable
function OleVariantArrayEnum(const Collection: OleVariant): IGetOleVariantEnum;

//------------------------------------------------------------------------------
//------------------------------------------------------------------------------
//------------------------------------------------------------------------------

implementation

uses
  Generics.Collections,
  JwaActiveX,
  VarUtils,
  ActiveX,
  Variants;

const
  // Impersonation Level Constants
  // http://msdn.microsoft.com/en-us/library/ms693790%28v=vs.85%29.aspx
  RPC_C_AUTHN_LEVEL_DEFAULT   = 0;
  RPC_C_IMP_LEVEL_ANONYMOUS   = 1;
  RPC_C_IMP_LEVEL_IDENTIFY    = 2;
  RPC_C_IMP_LEVEL_IMPERSONATE = 3;
  RPC_C_IMP_LEVEL_DELEGATE    = 4;

  // Authentication Service Constants
  // http://msdn.microsoft.com/en-us/library/ms692656%28v=vs.85%29.aspx
  RPC_C_AUTHN_WINNT      = 10;
  RPC_C_AUTHN_LEVEL_CALL = 3;
  RPC_C_AUTHN_DEFAULT    = Integer($FFFFFFFF);
  EOAC_NONE              = 0;

  // Authorization Constants
  // http://msdn.microsoft.com/en-us/library/ms690276%28v=vs.85%29.aspx
  RPC_C_AUTHZ_NONE       = 0;
  RPC_C_AUTHZ_NAME       = 1;
  RPC_C_AUTHZ_DCE        = 2;
  RPC_C_AUTHZ_DEFAULT    = Integer($FFFFFFFF);

  // Authentication-Level Constants
  // http://msdn.microsoft.com/en-us/library/aa373553%28v=vs.85%29.aspx
  RPC_C_AUTHN_LEVEL_PKT_PRIVACY   = 6;

  SEC_WINNT_AUTH_IDENTITY_ANSI    = 1;
  SEC_WINNT_AUTH_IDENTITY_UNICODE = 2;


//------------------------------------------------------------------------------
//
//      Error handling
//
//------------------------------------------------------------------------------
procedure WmiCheck(ErrorCode: HRESULT);
begin
  if (not Succeeded(ErrorCode)) then
    raise EWMIError.Create(ErrorCode);
end;

constructor EWMIError.Create(ErrorCode: HRESULT);
var
 pStatus: IWbemStatusCodeText;
 MessageText: WideString;
begin
  MessageText := '';

  if (Succeeded(CoCreateInstance(CLSID_WbemStatusCodeText, nil, CLSCTX_INPROC_SERVER, IID_IWbemStatusCodeText, pStatus))) then
  begin
    if (Succeeded(pStatus.GetErrorCodeText(ErrorCode, 0, 0, MessageText))) then
    begin
      // Trim trailing CR/LFs
{$WARN WIDECHAR_REDUCED OFF}
      while (Length(MessageText) > 0) and (MessageText[Length(MessageText)] in [#13, #10]) do
        SetLength(MessageText, Length(MessageText)-1);
{$WARN WIDECHAR_REDUCED DEFAULT}
    end else
      MessageText := '';
  end else
    MessageText := '';

  inherited Create(MessageText, ErrorCode, 0)
end;


//------------------------------------------------------------------------------
//
//      IWindowsManagementInternal
//
//------------------------------------------------------------------------------
type
  IWindowsManagementInternal = interface(IWindowsManagement)
    function GetServices: IWbemServices;
    function GetContext: IWbemContext;

    property Services: IWbemServices read GetServices;
    property Context: IWbemContext read GetContext;
  end;

//------------------------------------------------------------------------------
//
//      TWindowsManagementObject
//
//------------------------------------------------------------------------------
// Implements IWindowsManagementObject
//------------------------------------------------------------------------------
type
  TWindowsManagementObject = class abstract(TInterfacedObject, IWindowsManagementObject)
  private
    FConnection: IWindowsManagementInternal;
    FInstance: IWbemClassObject;
    FMethods: TDictionary<string, IWindowsManagementObjectMethod>; // Cache
    FKeyName: string; // Cached value
  protected
    class function TryGetQualifier(const AObject: IWbemClassObject; const AName: string; var Value: Variant): boolean;
    class function GetQualifier(const AObject: IWbemClassObject; const AName: string): Variant;

    // IWindowsManagementObject
    function GetNativeObject: IWbemClassObject;
    function GetValue(const AName: string): Variant;
    procedure SetValue(const AName: string; const AValue: Variant);
    function GetMethod(const AMethodName: string): IWindowsManagementObjectMethod;
    function GetIsClass: boolean;
    function GetIsInstance: boolean;
    function GetIsSingleton: boolean;
    function GetClassName: string;
    function GetKeyName: string;
    function GetPath: string;
    function GetRelativePath: string;
    function GetClassObject: IWindowsManagementClassObject; virtual; abstract;

    property NativeObject: IWbemClassObject read GetNativeObject;

    property IsClass: boolean read GetIsClass;
    property IsInstance: boolean read GetIsInstance;
    property IsSingleton: boolean read GetIsSingleton;

    property ClassObject: IWindowsManagementClassObject read GetClassObject;
{$WARN HIDING_MEMBER OFF}
    property ClassName: string read GetClassName;
{$WARN HIDING_MEMBER DEFAULT}
    property KeyName: string read GetKeyName;
    property Path: string read GetPath;
    property RelativePath: string read GetRelativePath;

    property Values[const AName: string]: Variant read GetValue write SetValue; default;
    property Methods[const AMethodName: string]: IWindowsManagementObjectMethod read GetMethod;
  public
    constructor Create(const AConnection: IWindowsManagementInternal; const AInstance: IWbemClassObject);
    destructor Destroy; override;
  end;


//------------------------------------------------------------------------------
//
//      TWindowsManagementObjectMethodResult
//
//------------------------------------------------------------------------------
// Implements IWindowsManagementObjectMethodResult
//------------------------------------------------------------------------------
type
  TWindowsManagementObjectMethodResult = class(TInterfacedObject, IWindowsManagementObjectMethodResult)
  private
    FOutParams: IWbemClassObject;
  protected
    function GetParamNames: string; // For debug use only
    // IWindowsNameSpaceMethodResult
    function GetParam(const AName: string): Variant;
  public
    constructor Create(const AOutParams: IWbemClassObject);
  end;

constructor TWindowsManagementObjectMethodResult.Create(const AOutParams: IWbemClassObject);
begin
  inherited Create;
  FOutParams := AOutParams;
end;

function TWindowsManagementObjectMethodResult.GetParamNames: string;
var
  ErrorCode: HRESULT;
  ParamName: WideString;
type
  POleVariant = ^OleVariant;
  PCIMTYPE = ^CIMTYPE;
begin
  Result := '';
  WmiCheck(FOutParams.BeginEnumeration(WBEM_FLAG_NONSYSTEM_ONLY));
  try

    while (True) do
    begin
      // Get the parameter name
      ErrorCode := FOutParams.Next(0, ParamName, POleVariant(nil)^, PCIMTYPE(nil)^, PInteger(nil)^);
      WmiCheck(ErrorCode);

      if (Result <> '') then
        Result := Result + #13;
      Result := Result + ParamName;

      if (ErrorCode = HRESULT(WBEM_S_NO_MORE_DATA)) then
        // No more params
        break;
    end;

  finally
    WmiCheck(FOutParams.EndEnumeration);
  end;
end;

function TWindowsManagementObjectMethodResult.GetParam(const AName: string): Variant;
var
  Value: OleVariant;
  Res: HResult;
begin
  Res := FOutParams.Get(PWideChar(AName), 0, Value, PInteger(nil)^, PInteger(nil)^);
  if (Succeeded(Res)) then
  begin
  try
    Result := Value;
  finally
    VariantClear(Value);
  end;
  end else
  begin
    if (DWORD(Res) = WBEM_E_NOT_FOUND) then
      Result := Unassigned
    else
      WmiCheck(Res);
  end;
end;

//------------------------------------------------------------------------------
//
//      TWindowsManagementObjectMethod
//
//------------------------------------------------------------------------------
// Note that is it possible to execute a method on both a class (i.e. a static method) and an instance.
//------------------------------------------------------------------------------
// Implements IWindowsManagementObjectMethod
//------------------------------------------------------------------------------
type
  TWindowsManagementObjectMethod = class(TInterfacedObject, IWindowsManagementObjectMethod)
  private
    FConnection: IWindowsManagementInternal;
    FInstance: IWindowsManagementObject;
    FInstancePath: string; // Cached path
    FMethodName: string;
    FParamsDefinition: IWbemClassObject;
    FParamsInstance: IWbemClassObject;
  protected
    // IWindowsNameSpaceMethod
    function GetIsStatic: boolean;
    procedure SetParam(const AName: string; const Value: Variant);
    function Execute: IWindowsManagementObjectMethodResult; overload;
    function Execute(const Params: array of Variant): IWindowsManagementObjectMethodResult; overload;
  public
    constructor Create(const AConnection: IWindowsManagementInternal; const AInstance: IWindowsManagementObject; const AMethodName: string);
  end;

constructor TWindowsManagementObjectMethod.Create(const AConnection: IWindowsManagementInternal; const AInstance: IWindowsManagementObject; const AMethodName: string);
var
  ResultParams: IWbemClassObject;
begin
  inherited Create;
  FConnection := AConnection;
  FInstance := AInstance;
  FMethodName := AMethodName;

  FParamsDefinition := nil;
  WmiCheck(FInstance.ClassObject.NativeObject.GetMethod(PWideChar(FMethodName), 0, FParamsDefinition, ResultParams));
end;

function TWindowsManagementObjectMethod.GetIsStatic: boolean;
var
  Value: Variant;
begin
  Result := (TWindowsManagementObject.TryGetQualifier(FParamsDefinition, 'Static', Value)) and (boolean(Value))
end;

function TWindowsManagementObjectMethod.Execute(const Params: array of Variant): IWindowsManagementObjectMethodResult;
type
  POleVariant = ^OleVariant;
  PCIMTYPE = ^CIMTYPE;
var
  OrderedParams: TDictionary<integer, string>;
  ErrorCode: HRESULT;
  ParamName: WideString;
  QualSet: IWbemQualifierSet;
  OleValue: OleVariant;
  Value: Variant;
  Index: integer;
  Name: string;
  s: string;
begin
  if (Length(Params) > 0) then
  begin
    if (FParamsDefinition = nil) then
      raise Exception.Create('Method has no parameters');

    OrderedParams := TDictionary<integer, string>.Create;
    try

      WmiCheck(FParamsDefinition.BeginEnumeration(WBEM_FLAG_NONSYSTEM_ONLY));
      try

        while (True) do
        begin
          // Get the parameter name
          ErrorCode := FParamsDefinition.Next(0, ParamName, POleVariant(nil)^, PCIMTYPE(nil)^, PInteger(nil)^);
          WmiCheck(ErrorCode);

          if (ErrorCode = HRESULT(WBEM_S_NO_MORE_DATA)) then
            // No more params
            break;

          // Get the parameter qualifier set
          WmiCheck(FParamsDefinition.GetPropertyQualifierSet(PChar(ParamName), QualSet));

          // Get the parameter order
          WmiCheck(QualSet.Get('ID', 0, OleValue, PInteger(nil)^));
          try
            // Save order/name pair
            OrderedParams.Add(integer(OleValue), ParamName);
          finally
            VariantClear(OleValue);
          end;
        end;

      finally
        WmiCheck(FParamsDefinition.EndEnumeration);
      end;

      // Assign parameter values in order
      for Index := 0 to Length(Params)-1 do
      begin
        if (not OrderedParams.TryGetValue(Index, Name)) then
          break;

        Value := Params[Index];

        // Work around "The parameter is incorrect" when passing empty string values
        if (VarIsStr(Value)) then
        begin
          s := Value;
          if (s = '') then
            Value := #0;
        end;

        SetParam(Name, Value);
      end;

    finally
      OrderedParams.Free;
    end;
  end;

  Result := Execute;
end;

function TWindowsManagementObjectMethod.Execute: IWindowsManagementObjectMethodResult;
var
  ResultParams: IWbemClassObject;
  CallResult: IWbemCallResult;
begin
  if (FInstancePath = '') then
    FInstancePath := FInstance.RelativePath;

  ResultParams := nil;
  WmiCheck(FConnection.Services.ExecMethod(FInstancePath, FMethodName, 0, nil, FParamsInstance, ResultParams, CallResult));

  if (ResultParams <> nil) then
    Result := TWindowsManagementObjectMethodResult.Create(ResultParams)
  else
    Result := nil;
end;

procedure TWindowsManagementObjectMethod.SetParam(const AName: string; const Value: Variant);
var
 Val: OleVariant;
begin
  if (FParamsDefinition = nil) then
    raise Exception.Create('Method has no parameters');

  if (FParamsInstance = nil) then
    WmiCheck(FParamsDefinition.SpawnInstance(0, FParamsInstance));

  Val := Value;

  WmiCheck(FParamsInstance.Put(PWideChar(AName), 0, @Val, 0));
end;

//------------------------------------------------------------------------------
//
//      TWindowsManagementObject
//
//------------------------------------------------------------------------------
// Implements IWindowsManagementObject
//------------------------------------------------------------------------------
constructor TWindowsManagementObject.Create(const AConnection: IWindowsManagementInternal; const AInstance: IWbemClassObject);
begin
  inherited Create;
  FConnection := AConnection;
  FInstance := AInstance;
end;

destructor TWindowsManagementObject.Destroy;
begin
  FreeAndNil(FMethods);
  inherited;
end;

function TWindowsManagementObject.GetNativeObject: IWbemClassObject;
begin
  Result := FInstance;
end;

function TWindowsManagementObject.GetClassName: string;
begin
  Result := VarToStr(Values['__CLASS']);
end;

function TWindowsManagementObject.GetIsClass: boolean;
begin
  Result := (Values['__Genus'] = WBEM_GENUS_CLASS);
end;

function TWindowsManagementObject.GetIsInstance: boolean;
begin
  Result := (Values['__Genus'] = WBEM_GENUS_INSTANCE);
end;

function TWindowsManagementObject.GetIsSingleton: boolean;
var
  Value: Variant;
begin
  Result := (TWindowsManagementObject.TryGetQualifier(ClassObject.NativeObject, 'Singleton', Value)) and (boolean(Value));
end;

function TWindowsManagementObject.GetKeyName: string;
var
  Names: PSafeArray;
  LowIndex, HighIndex: integer;
  Value: WideString;
begin
  if (FKeyName = '') then
  begin
    // Get name(s) of Key properties (hopefully only one)
    Names := nil;
    WmiCheck(NativeObject.GetNames(nil, WBEM_FLAG_KEYS_ONLY, nil, JwaActiveX.PSafeArray(Names)));
    try
      if (Names.cDims <> 1) then
        raise Exception.CreateFmt('Invalid dimension count for Key property list: %d', [Names.cDims]);

      SafeArrayGetLBound(Names, 1, LowIndex);
      SafeArrayGetUBound(Names, 1, HighIndex);

      if (LowIndex <> HighIndex) then
        raise Exception.CreateFmt('Too many Key properties: %d', [HighIndex-LowIndex+1]);

      SafeArrayGetElement(Names, LowIndex, Value);
    finally
      SafeArrayDestroy(Names);
    end;

    FKeyName := Value;
  end;

  Result := FKeyName;
end;

function TWindowsManagementObject.GetMethod(const AMethodName: string): IWindowsManagementObjectMethod;
begin
  // Try to resolve from cache
  if (FMethods <> nil) and (FMethods.TryGetValue(AnsiUpperCase(AMethodName), Result)) then
    exit;

  Result := TWindowsManagementObjectMethod.Create(FConnection, Self, AMethodName);

  if (FMethods = nil) then
    FMethods := TDictionary<string, IWindowsManagementObjectMethod>.Create;

  // Add to cache
  FMethods.Add(AnsiUpperCase(AMethodName), Result);
end;

function TWindowsManagementObject.GetPath: string;
begin
  Result := VarToStr(Values['__PATH']);
end;

class function TWindowsManagementObject.TryGetQualifier(const AObject: IWbemClassObject; const AName: string; var Value: Variant): boolean;
var
  QualifierSet: IWbemQualifierSet;
  Val: OleVariant;
  Res: HRESULT;
begin
  Value := Unassigned;
  Result := False;

  WmiCheck(AObject.GetQualifierSet(QualifierSet));

  Res := QualifierSet.Get(PWideChar(AName), 0, Val, PInteger(nil)^);
  try

    if (DWORD(Res) = WBEM_NO_ERROR) then
    begin
      Value := Val;
      Result := True;
    end else
    if (DWORD(Res) <> WBEM_E_NOT_FOUND) then
      WmiCheck(Res);

  finally
    VariantClear(Val);
  end;
end;

class function TWindowsManagementObject.GetQualifier(const AObject: IWbemClassObject; const AName: string): Variant;
begin
  if (not TryGetQualifier(AObject, AName, Result)) then
    Result := Unassigned;
end;

function TWindowsManagementObject.GetRelativePath: string;
begin
  Result := VarToStr(Values['__RELPATH']);
end;

function TWindowsManagementObject.GetValue(const AName: string): Variant;
var
 Val: OleVariant;
begin
  WmiCheck(FInstance.Get(PWideChar(AName), 0, Val, PInteger(nil)^, PInteger(nil)^));
  try
    Result := Val;
  finally
    VariantClear(Val);
  end;
end;

procedure TWindowsManagementObject.SetValue(const AName: string; const AValue: Variant);
var
 Val: OleVariant;
begin
  Val := AValue;

  WmiCheck(FInstance.Put(PWideChar(AName), 0, @Val, 0));
end;


//------------------------------------------------------------------------------
//
//      TWindowsManagementInstanceObject
//
//------------------------------------------------------------------------------
// Wraps an instance
//------------------------------------------------------------------------------
type
  TWindowsManagementInstanceObject = class(TWindowsManagementObject, IWindowsManagementInstanceObject)
  private
    FClassObject: IWindowsManagementClassObject;
  protected
    // IWindowsManagementObject
    function GetClassObject: IWindowsManagementClassObject; override;
    // IWindowsManagementInstanceObject
    function GetInstanceName: string;
  public
    constructor Create(const AConnection: IWindowsManagementInternal; const AInstance: IWbemClassObject); overload;
    constructor Create(const AConnection: IWindowsManagementInternal; const AClassObject: IWindowsManagementClassObject; const AInstance: IWbemClassObject); overload;
  end;

constructor TWindowsManagementInstanceObject.Create(const AConnection: IWindowsManagementInternal; const AClassObject: IWindowsManagementClassObject; const AInstance: IWbemClassObject);
begin
  Create(AConnection, AInstance);

  FClassObject := AClassObject;
end;

constructor TWindowsManagementInstanceObject.Create(const AConnection: IWindowsManagementInternal; const AInstance: IWbemClassObject);
begin
  inherited Create(AConnection, AInstance);
end;

function TWindowsManagementInstanceObject.GetClassObject: IWindowsManagementClassObject;
begin
  if (FClassObject = nil) then
    FClassObject := FConnection.Objects[ClassName];

  Result := FClassObject;
end;

function TWindowsManagementInstanceObject.GetInstanceName: string;
begin
  Result := Values[KeyName];
end;

//------------------------------------------------------------------------------
//
//      TWindowsManagementInstanceEnumerator
//
//------------------------------------------------------------------------------
type
  TWindowsManagementInstanceEnumerator = class(TInterfacedObject, IWindowsManagementInstanceEnumerator)
  private
    FConnection: IWindowsManagementInternal;
    FClassObject: IWindowsManagementClassObject;
    FEnum: IEnumWbemClassObject;
    FInstance: IWbemClassObject;
  protected
    function GetCurrent: IWindowsManagementInstanceObject;
    property Current: IWindowsManagementInstanceObject read GetCurrent;
    function MoveNext: Boolean;
  public
    constructor Create(const AConnection: IWindowsManagementInternal; const AEnum: IEnumWbemClassObject); overload;
    constructor Create(const AConnection: IWindowsManagementInternal; const AClassObject: IWindowsManagementClassObject; const AEnum: IEnumWbemClassObject); overload;
  end;

constructor TWindowsManagementInstanceEnumerator.Create(const AConnection: IWindowsManagementInternal; const AEnum: IEnumWbemClassObject);
begin
  inherited Create;

  FConnection := AConnection;
  FEnum := AEnum;
end;

constructor TWindowsManagementInstanceEnumerator.Create(const AConnection: IWindowsManagementInternal; const AClassObject: IWindowsManagementClassObject; const AEnum: IEnumWbemClassObject);
begin
  Create(AConnection, AEnum);

  FClassObject := AClassObject;
end;

function TWindowsManagementInstanceEnumerator.GetCurrent: IWindowsManagementInstanceObject;
begin
  if (FInstance <> nil) then
    Result := TWindowsManagementInstanceObject.Create(FConnection, FClassObject, FInstance)
  else
    Result := nil;
end;

function TWindowsManagementInstanceEnumerator.MoveNext: Boolean;
var
  Res: HRESULT;
  Count: Cardinal;
begin
  FInstance := nil;

  Res := FEnum.Next(integer(WBEM_INFINITE), 1, FInstance, Count);

  Result := (Res = WBEM_S_NO_ERROR);

  if (not Result) then
    WmiCheck(Res);
end;


//------------------------------------------------------------------------------
//
//      TCustomWindowsManagementInstanceList
//
//------------------------------------------------------------------------------
type
  TCustomWindowsManagementInstanceList = class(TInterfacedObject, IWindowsManagementSimpleInstanceList)
  strict private
    FConnection: IWindowsManagementInternal;
  protected
    property Connection: IWindowsManagementInternal read FConnection;

    // IWindowsNameSpaceInstanceList
    function GetFirst: IWindowsManagementInstanceObject;
    property First: IWindowsManagementInstanceObject read GetFirst;
  public
    constructor Create(const AConnection: IWindowsManagementInternal);

    function GetEnumerator: IWindowsManagementInstanceEnumerator; virtual; abstract;
  end;

constructor TCustomWindowsManagementInstanceList.Create(const AConnection: IWindowsManagementInternal);
begin
  inherited Create;
  FConnection := AConnection;
end;

function TCustomWindowsManagementInstanceList.GetFirst: IWindowsManagementInstanceObject;
var
  Instance: IWindowsManagementInstanceObject;
begin
  Result := nil;

  for Instance in Self do
  begin
    Result := Instance;
    break;
  end;
end;

//------------------------------------------------------------------------------
//
//      TWindowsManagementInstanceList
//
//------------------------------------------------------------------------------
// A list that wraps an IEnumWbemClassObject. Only contains the enumerator.
//------------------------------------------------------------------------------
type
  TWindowsManagementInstanceList = class(TCustomWindowsManagementInstanceList)
  strict private
    FEnum: IEnumWbemClassObject;
  protected
  public
    constructor Create(const AConnection: IWindowsManagementInternal; const AEnum: IEnumWbemClassObject);

    function GetEnumerator: IWindowsManagementInstanceEnumerator; override;
  end;

constructor TWindowsManagementInstanceList.Create(const AConnection: IWindowsManagementInternal; const AEnum: IEnumWbemClassObject);
begin
  inherited Create(AConnection);
  FEnum := AEnum;
end;

function TWindowsManagementInstanceList.GetEnumerator: IWindowsManagementInstanceEnumerator;
begin
  Result := TWindowsManagementInstanceEnumerator.Create(Connection, FEnum);
end;

//------------------------------------------------------------------------------
//
//      TWindowsManagementObjectList
//
//------------------------------------------------------------------------------
// A list that wraps a scope.
//------------------------------------------------------------------------------
type
  TWindowsManagementObjectList = class(TCustomWindowsManagementInstanceList, IWindowsManagementInstanceList)
  private
    FClassObject: IWindowsManagementClassObject;
    FItems: TDictionary<string, IWindowsManagementInstanceObject>; // Cache
    FPath: string;
  protected
    function GetItem(const AInstanceName: string): IWindowsManagementInstanceObject;
  public
    constructor Create(const AConnection: IWindowsManagementInternal; const AClassObject: IWindowsManagementClassObject);
    destructor Destroy; override;

    property Items[const AInstanceName: string]: IWindowsManagementInstanceObject read GetItem; default;

    function GetEnumerator: IWindowsManagementInstanceEnumerator; override;
  end;

constructor TWindowsManagementObjectList.Create(const AConnection: IWindowsManagementInternal; const AClassObject: IWindowsManagementClassObject);
begin
  inherited Create(AConnection);
  FClassObject := AClassObject;

  if (FClassObject <> nil) then
    FPath := FClassObject.RelativePath;
end;

destructor TWindowsManagementObjectList.Destroy;
begin
  FreeAndNil(FItems);

  inherited;
end;

function TWindowsManagementObjectList.GetEnumerator: IWindowsManagementInstanceEnumerator;
var
  Enum: IEnumWbemClassObject;
begin
  WmiCheck(Connection.Services.CreateInstanceEnum(FPath, WBEM_FLAG_FORWARD_ONLY or WBEM_FLAG_RETURN_IMMEDIATELY, Connection.Context, Enum));

  Result := TWindowsManagementInstanceEnumerator.Create(Connection, FClassObject, Enum);
end;

function TWindowsManagementObjectList.GetItem(const AInstanceName: string): IWindowsManagementInstanceObject;
var
  Path: string;
  Instance: IWbemClassObject;
  CallResult: IWbemCallResult;
begin
  // Try to resolve from cache
  if (FItems <> nil) and (FItems.TryGetValue(AnsiUpperCase(AInstanceName), Result)) then
      exit;

  Path := Format('%s.%s=''%s''', [FClassObject.ClassName, FClassObject.KeyName, AInstanceName]);

  WmiCheck(Connection.Services.GetObject(Path, WBEM_FLAG_RETURN_WBEM_COMPLETE, Connection.Context, Instance, CallResult));

  if (Instance <> nil) then
  begin
    Result := TWindowsManagementInstanceObject.Create(Connection, FClassObject, Instance);

    if (FItems = nil) then
      FItems := TDictionary<string, IWindowsManagementInstanceObject>.Create;

    // Add to cache
    FItems.Add(AnsiUpperCase(AInstanceName), Result);
  end else
    Result := nil;
end;

//------------------------------------------------------------------------------
//
//      TWindowsManagementClassObject
//
//------------------------------------------------------------------------------
// Wraps a class
//------------------------------------------------------------------------------
type
  TWindowsManagementClassObject = class(TWindowsManagementObject, IWindowsManagementClassObject)
  private
    FInstances: IWindowsManagementInstanceList;
  protected
    // IWindowsManagementObject
    function GetClassObject: IWindowsManagementClassObject; override;

    // IWindowsManagementClassObject
    function GetInstances: IWindowsManagementInstanceList;
    function GetInstance: IWindowsManagementInstanceObject;

    property Instance: IWindowsManagementInstanceObject read GetInstance;
    property Instances: IWindowsManagementInstanceList read GetInstances;
  public
  end;

function TWindowsManagementClassObject.GetClassObject: IWindowsManagementClassObject;
begin
  Result := Self;
end;

function TWindowsManagementClassObject.GetInstance: IWindowsManagementInstanceObject;
begin
  if (IsSingleton) then
    Result := Instances.First
  else
    Result := nil;
end;

function TWindowsManagementClassObject.GetInstances: IWindowsManagementInstanceList;
begin
  if (FInstances = nil) then
    FInstances := TWindowsManagementObjectList.Create(FConnection, Self);

  Result := FInstances;
end;

//------------------------------------------------------------------------------
//
//      TWindowsManagementClassEnumerator
//
//------------------------------------------------------------------------------
// A class enumerator
//------------------------------------------------------------------------------
type
  TWindowsManagementClassEnumerator = class(TInterfacedObject, IWindowsManagementClassEnumerator)
  private
    FConnection: IWindowsManagementInternal;
    FEnum: IEnumWbemClassObject;
    FInstance: IWbemClassObject;
  protected
    function GetCurrent: IWindowsManagementClassObject;
    property Current: IWindowsManagementClassObject read GetCurrent;
    function MoveNext: Boolean;
  public
    constructor Create(const AConnection: IWindowsManagementInternal; const AEnum: IEnumWbemClassObject);
  end;

constructor TWindowsManagementClassEnumerator.Create(const AConnection: IWindowsManagementInternal; const AEnum: IEnumWbemClassObject);
begin
  inherited Create;

  FConnection := AConnection;
  FEnum := AEnum;
end;

function TWindowsManagementClassEnumerator.GetCurrent: IWindowsManagementClassObject;
begin
  if (FInstance <> nil) then
    Result := TWindowsManagementClassObject.Create(FConnection, FInstance)
  else
    Result := nil;
end;

function TWindowsManagementClassEnumerator.MoveNext: Boolean;
var
  Res: HRESULT;
  Count: Cardinal;
begin
  FInstance := nil;

  Res := FEnum.Next(integer(WBEM_INFINITE), 1, FInstance, Count);

  Result := (Res = WBEM_S_NO_ERROR);

  if (not Result) then
    WmiCheck(Res);
end;


//------------------------------------------------------------------------------
//
//      TWindowsManagementClassList
//
//------------------------------------------------------------------------------
// A list of classes with random access
//------------------------------------------------------------------------------
type
  TWindowsManagementClassList = class(TInterfacedObject, IWindowsManagementClassList)
  strict private
    FConnection: IWindowsManagementInternal;
    FItems: TDictionary<string, IWindowsManagementClassObject>; // Cache
  protected
    function GetItem(const APath: string): IWindowsManagementClassObject;
  public
    constructor Create(const AConnection: IWindowsManagementInternal);
    destructor Destroy; override;

    property Items[const APath: string]: IWindowsManagementClassObject read GetItem; default;

    function GetEnumerator: IWindowsManagementClassEnumerator;
  end;

constructor TWindowsManagementClassList.Create(const AConnection: IWindowsManagementInternal);
begin
  inherited Create;
  FConnection := AConnection;
end;

destructor TWindowsManagementClassList.Destroy;
begin
  FreeAndNil(FItems);
  inherited;
end;

function TWindowsManagementClassList.GetEnumerator: IWindowsManagementClassEnumerator;
var
  Enum: IEnumWbemClassObject;
begin
  WmiCheck(FConnection.Services.CreateClassEnum('', WBEM_FLAG_FORWARD_ONLY or WBEM_FLAG_RETURN_IMMEDIATELY, FConnection.Context, Enum));

  Result := TWindowsManagementClassEnumerator.Create(FConnection, Enum);
end;

function TWindowsManagementClassList.GetItem(const APath: string): IWindowsManagementClassObject;
var
  Instance: IWbemClassObject;
  CallResult: IWbemCallResult;
begin
  // Try to resolve from cache
  if (FItems <> nil) and (FItems.TryGetValue(AnsiUpperCase(APath), Result)) then
    exit;

  WmiCheck(FConnection.Services.GetObject(APath, WBEM_FLAG_RETURN_WBEM_COMPLETE, FConnection.Context, Instance, CallResult));

  if (Instance <> nil) then
  begin
    Result := TWindowsManagementClassObject.Create(FConnection, Instance);

    if (FItems = nil) then
      FItems := TDictionary<string, IWindowsManagementClassObject>.Create;

    // Add item to cache
    FItems.Add(AnsiUpperCase(APath), Result);
  end else
    Result := nil;
end;


//------------------------------------------------------------------------------
//
//      TWindowsManagement
//
//------------------------------------------------------------------------------
// Wraps a WMI connection and scope/namespace
//------------------------------------------------------------------------------
// COAUTHIDENTITY Structure
// http://msdn.microsoft.com/en-us/library/ms693358%28v=vs.85%29.aspx
type
  PCOAUTHIDENTITY    = ^TCOAUTHIDENTITY;
  _COAUTHIDENTITY    = record
    User           : PChar;
    UserLength     : ULONG;
    Domain         : PChar;
    DomainLength   : ULONG;
    Password       : PChar;
    PassWordLength : ULONG;
    Flags          : ULONG;
  end;

  COAUTHIDENTITY      = _COAUTHIDENTITY;
  TCOAUTHIDENTITY     = _COAUTHIDENTITY;

type
  TWindowsManagement = class(TInterfacedObject, IWindowsManagement, IWindowsManagementInternal)
  strict private
    FScope: string;
    FComputerName: string;
    FUserName: string;
    FPassword: string;
    FLocale: string;
    FAuthInfo: TCOAUTHIDENTITY;
    FContext: IWbemContext;

    FLocalConnection: Boolean;
    FLocator: IWbemLocator;
    FServices: IWbemServices;
    FUnsecuredApartment: IUnsecuredApartment;

    FObjects: IWindowsManagementClassList; // Cached value

  strict protected
    // IWindowsManagement
    function GetObjects: IWindowsManagementClassList;
    function ExecuteQuery(const Query: string): IWindowsManagementSimpleInstanceList;
    property Objects: IWindowsManagementClassList read GetObjects;

    // IWindowsManagementInternal
    function GetServices: IWbemServices;
    function GetContext: IWbemContext;

  protected
    property Scope: string read FScope;
    property UserName: string read FUserName;
    property Password: string read FPassword;
    property ComputerName: string read FComputerName;
    property Locale: string read FLocale;
    property LocalConnection: Boolean read FLocalConnection;
    property Services: IWbemServices read FServices;
    property Context: IWbemContext read FContext write FContext;

  public
    constructor Create(const AScope: string = ''; const AComputerName: string = 'localhost'; const AUserName: string = ''; const APassword: string = ''; const AContext: IWbemContext = nil);
    destructor Destroy; override;
  end;

//------------------------------------------------------------------------------

constructor TWindowsManagement.Create(const AScope, AComputerName, AUserName, APassword: string; const AContext: IWbemContext);
var
  Path: string;
begin
  inherited Create;

  FScope := AScope;
  if (Length(FScope) > 0) and (FScope[1] <> '\') then
    Insert('\', FScope, 1);

  FUserName := AUserName;
  FPassword := APassword;
  if (AComputerName <> '') then
    FComputerName := AComputerName
  else
    FComputerName := 'localhost';
  FContext := AContext;

  FLocale := '';

  ZeroMemory(@FAuthInfo, 0);
  FAuthInfo.User := PChar(FUserName);
  FAuthInfo.UserLength := Length(FUserName);
  FAuthInfo.Domain := '';
  FAuthInfo.DomainLength := 0;
  FAuthInfo.Password := PChar(FPassword);
  FAuthInfo.PasswordLength := Length(FPassword);
  FAuthInfo.Flags := SEC_WINNT_AUTH_IDENTITY_UNICODE;

  FLocalConnection := (SameText(FComputerName, 'localhost'));

  OleCheck(CoCreateInstance(CLSID_WbemLocator, nil, CLSCTX_INPROC_SERVER, IID_IWbemLocator, FLocator));

  Path := Format('\\%s%s', [FComputerName, FScope]);

  WmiCheck(FLocator.ConnectServer(Path, FUserName, FPassword, FLocale,  WBEM_FLAG_CONNECT_USE_MAX_WAIT, '', FContext, FServices));

  // Set security levels on a WMI connection
  if (FLocalConnection) then
    OleCheck(CoSetProxyBlanket(FServices, RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE, nil, RPC_C_AUTHN_LEVEL_CALL, RPC_C_IMP_LEVEL_IMPERSONATE, nil, EOAC_NONE))
  else
    OleCheck(CoSetProxyBlanket(FServices,  RPC_C_AUTHN_DEFAULT, RPC_C_AUTHZ_DEFAULT, PWideChar(Format('\\%s', [FComputerName])), RPC_C_AUTHN_LEVEL_PKT_PRIVACY, RPC_C_IMP_LEVEL_IMPERSONATE, @FAuthInfo, EOAC_NONE));

  OleCheck(CoCreateInstance(CLSID_UnsecuredApartment, nil, CLSCTX_LOCAL_SERVER, IID_IUnsecuredApartment, FUnsecuredApartment));
end;

destructor TWindowsManagement.Destroy;
begin

  inherited;
end;

function TWindowsManagement.ExecuteQuery(const Query: string): IWindowsManagementSimpleInstanceList;
var
  Enum: IEnumWbemClassObject;
begin
  if (FServices = nil) then
    raise Exception.Create('TWindowsNameSpace has not been connected');

  WmiCheck(FServices.ExecQuery('WQL', Query, WBEM_FLAG_FORWARD_ONLY or WBEM_FLAG_RETURN_IMMEDIATELY, FContext, Enum));

  Result := TWindowsManagementInstanceList.Create(Self, Enum);
end;

function TWindowsManagement.GetContext: IWbemContext;
begin
  Result := FContext;
end;

function TWindowsManagement.GetObjects: IWindowsManagementClassList;
begin
  if (FObjects = nil) then
    FObjects := TWindowsManagementClassList.Create(Self);

  Result := FObjects;
end;

function TWindowsManagement.GetServices: IWbemServices;
begin
  Result := FServices;
end;


//------------------------------------------------------------------------------
//
//      Main entry point/factory
//
//------------------------------------------------------------------------------
function ConnectWindowsManagement(const AScope, AComputerName, AUserName, APassword: string; const AContext: IWbemContext): IWindowsManagement;
begin
  Result := TWindowsManagement.Create(AScope, AComputerName, AUserName, APassword, AContext);
end;

procedure WindowsManagementInitializeSecurity(LocalConnection: boolean);
begin
  if LocalConnection then
    OleCheck(CoInitializeSecurity(nil, -1, nil, nil, RPC_C_AUTHN_LEVEL_DEFAULT, RPC_C_IMP_LEVEL_IMPERSONATE, nil, EOAC_NONE, nil))
  else
    OleCheck(CoInitializeSecurity(nil, -1, nil, nil, RPC_C_AUTHN_LEVEL_DEFAULT, RPC_C_IMP_LEVEL_IDENTIFY, nil, EOAC_NONE, nil));
end;

//------------------------------------------------------------------------------
//
//      Variant array enumerator
//
//------------------------------------------------------------------------------
type
  TOleVariantEnum = class(TInterfacedObject, IOleVariantEnum, IGetOleVariantEnum)
  private
    FCurrent: OLEVariant;
    FEnum: IEnumVARIANT;
  public
    function GetEnumerator: IOleVariantEnum;
    constructor Create(const Collection: OLEVariant);
    function GetCurrent: OLEVariant;
    function MoveNext: Boolean;
    property Current: OLEVariant read GetCurrent;
  end;

  TOleVariantArrayEnum = class(TInterfacedObject, IOleVariantEnum, IGetOleVariantEnum)
  private
    FCollection: OLEVariant;
    FIndex: Integer;
    FLowBound: Integer;
    FHighBound: Integer;
  public
    function GetEnumerator: IOleVariantEnum;
    constructor Create(const Collection: OLEVariant);
    function  GetCurrent: OLEVariant;
    function  MoveNext: Boolean;
    property  Current: OLEVariant read GetCurrent;
  end;

constructor TOleVariantEnum.Create(const Collection: OLEVariant);
begin
  inherited Create;
  FEnum := IUnknown(Collection._NewEnum) As IEnumVARIANT;
end;

function TOleVariantEnum.GetCurrent: OLEVariant;
begin
  Result := FCurrent;
end;

function TOleVariantEnum.GetEnumerator: IOleVariantEnum;
begin
  Result := Self;
end;

function TOleVariantEnum.MoveNext: Boolean;
var
  Count: LongWord;
begin
  FCurrent := Unassigned;//avoid memory leaks
  Result := (FEnum.Next(1, FCurrent, Count) = S_OK);
end;

{ TOleVariantArrayEnum }

constructor TOleVariantArrayEnum.Create(const Collection: OLEVariant);
begin
  inherited Create;
  FCollection := Collection;
  FLowBound := VarArrayLowBound(FCollection, 1);
  FHighBound:= VarArrayHighBound(FCollection, 1);
  FIndex := FLowBound-1;
end;

function TOleVariantArrayEnum.GetCurrent: OLEVariant;
begin
  Result := FCollection[FIndex];
end;

function TOleVariantArrayEnum.GetEnumerator: IOleVariantEnum;
begin
  Result := Self;
end;

function TOleVariantArrayEnum.MoveNext: Boolean;
begin
  Result := (FIndex < FHighBound);
  if Result then
    Inc(FIndex);
end;

//------------------------------------------------------------------------------

function OleVariantEnum(const Collection: OleVariant): IGetOleVariantEnum;
begin
  Result := TOleVariantEnum.Create(Collection);
end;

function OleVariantArrayEnum(const Collection: OleVariant): IGetOleVariantEnum;
begin
  Result := TOleVariantArrayEnum.Create(Collection);
end;

//------------------------------------------------------------------------------
//
//      IWindowsManagementObjectMethodResult custom variant
//
//------------------------------------------------------------------------------
type
  TWindowsManagementObjectMethodResultVarData = packed record
    { Var type, will be assigned at runtime }
    VType: TVarType;
    { Reserved stuff }
    Reserved1, Reserved2, Reserved3: Word;
    { A reference to the enclosed object }
    WmiObject: IWindowsManagementObjectMethodResult;
    { Reserved stuff }
    Reserved4: LongWord;
  end;

  TWindowsManagementObjectMethodResultVariantType = class(TInvokeableVariantType)
  private
  public
    procedure Clear(var V: TVarData); override;
    procedure Copy(var Dest: TVarData; const Source: TVarData; const Indirect: Boolean); override;
    function GetProperty(var Dest: TVarData; const V: TVarData; const Name: string): Boolean; override;
  end;

var
  FWindowsManagementObjectMethodResultVariantType: TWindowsManagementObjectMethodResultVariantType;

procedure TWindowsManagementObjectMethodResultVariantType.Clear(var V: TVarData);
begin
  V.VType := varEmpty;
  TWindowsManagementObjectMethodResultVarData(V).WmiObject := nil;
end;

procedure TWindowsManagementObjectMethodResultVariantType.Copy(var Dest: TVarData; const Source: TVarData; const Indirect: Boolean);
begin
  if Indirect and VarDataIsByRef(Source) then
    VarDataCopyNoInd(Dest, Source)
  else
  begin
    TWindowsManagementObjectMethodResultVarData(Dest).VType := VarType;
    TWindowsManagementObjectMethodResultVarData(Dest).WmiObject := TWindowsManagementObjectMethodResultVarData(Source).WmiObject;
  end;
end;

function TWindowsManagementObjectMethodResultVariantType.GetProperty(var Dest: TVarData; const V: TVarData; const Name: string): Boolean;
begin
  Variant(Dest) := TWindowsManagementObjectMethodResultVarData(V).WmiObject[Name];
  Result := True;
end;


function WindowsManagementObjectMethodResult(const MethodResult: IWindowsManagementObjectMethodResult): Variant;
begin
  VarClear(Result);

  { Assign the new variant the var type that was allocated for us }
  TWindowsManagementObjectMethodResultVarData(Result).VType := FWindowsManagementObjectMethodResultVariantType.VarType;

  TWindowsManagementObjectMethodResultVarData(Result).WmiObject := MethodResult;
end;

//------------------------------------------------------------------------------
//
//      IWindowsManagementObject custom variant
//
//------------------------------------------------------------------------------
type
  TWindowsManagementObjectVarData = packed record
    { Var type, will be assigned at runtime }
    VType: TVarType;
    { Reserved stuff }
    Reserved1, Reserved2, Reserved3: Word;
    { A reference to the enclosed object }
    WmiObject: IWindowsManagementObject;
    { Reserved stuff }
    Reserved4: LongWord;
  end;

  TWindowsManagementObjectVariantType = class(TInvokeableVariantType)
  private
  public
    procedure Clear(var V: TVarData); override;
    procedure Copy(var Dest: TVarData; const Source: TVarData; const Indirect: Boolean); override;
    procedure Cast(var Dest: TVarData; const Source: TVarData); override;
    procedure CastTo(var Dest: TVarData; const Source: TVarData; const AVarType: TVarType); override;
    function GetProperty(var Dest: TVarData; const V: TVarData; const Name: string): Boolean; override;
    function SetProperty(const V: TVarData; const Name: string; const Value: TVarData): Boolean; override;
    function DoFunction(var Dest: TVarData; const V: TVarData; const Name: string; const Arguments: TVarDataArray): Boolean; override;
    function DoProcedure(const V: TVarData; const Name: string; const Arguments: TVarDataArray): Boolean; override;
  end;

var
  FWindowsManagementVariantType: TWindowsManagementObjectVariantType = nil;

procedure TWindowsManagementObjectVariantType.Cast(var Dest: TVarData; const Source: TVarData);
begin
  inherited;

end;

procedure TWindowsManagementObjectVariantType.CastTo(var Dest: TVarData; const Source: TVarData; const AVarType: TVarType);
begin
  if (Source.VType = VarType) then
  begin
    inherited;

  end else
    inherited;
end;

procedure TWindowsManagementObjectVariantType.Clear(var V: TVarData);
begin
  V.VType := varEmpty;
  TWindowsManagementObjectVarData(V).WmiObject := nil;
end;

procedure TWindowsManagementObjectVariantType.Copy(var Dest: TVarData; const Source: TVarData; const Indirect: Boolean);
begin
  if Indirect and VarDataIsByRef(Source) then
    VarDataCopyNoInd(Dest, Source)
  else
  begin
    TWindowsManagementObjectVarData(Dest).VType := VarType;
    TWindowsManagementObjectVarData(Dest).WmiObject := TWindowsManagementObjectVarData(Source).WmiObject;
  end;
end;

function TWindowsManagementObjectVariantType.DoFunction(var Dest: TVarData; const V: TVarData; const Name: string; const Arguments: TVarDataArray): Boolean;
var
  Method: IWindowsManagementObjectMethod;
  i: integer;
  Params: array of Variant;
  MethodResult: IWindowsManagementObjectMethodResult;
begin
  Method := TWindowsManagementObjectVarData(V).WmiObject.Methods[Name];

  SetLength(Params, Length(Arguments));
  for i := 0 to Length(Arguments)-1 do
    Params[i] := Variant(Arguments[i]);

  MethodResult := Method.Execute(Params);

  VarUtils.VariantInit(Dest);
  TWindowsManagementObjectMethodResultVarData(Dest).VType := FWindowsManagementObjectMethodResultVariantType.VarType;
  TWindowsManagementObjectMethodResultVarData(Dest).WmiObject := MethodResult;

  Result := True;
end;

function TWindowsManagementObjectVariantType.DoProcedure(const V: TVarData; const Name: string; const Arguments: TVarDataArray): Boolean;
var
  Method: IWindowsManagementObjectMethod;
  i: integer;
  Params: array of Variant;
begin
  Method := TWindowsManagementObjectVarData(V).WmiObject.Methods[Name];

  SetLength(Params, Length(Arguments));
  for i := 0 to Length(Arguments)-1 do
    Params[i] := Variant(Arguments[i]);

  Method.Execute(Params);

  Result := True;
end;

function TWindowsManagementObjectVariantType.GetProperty(var Dest: TVarData; const V: TVarData; const Name: string): Boolean;
begin
  Dest := TVarData(TWindowsManagementObjectVarData(V).WmiObject[Name]);
  Result := True;
end;

function TWindowsManagementObjectVariantType.SetProperty(const V: TVarData; const Name: string; const Value: TVarData): Boolean;
begin
  TWindowsManagementObjectVarData(V).WmiObject[Name] := Variant(Value);
  Result := True;
end;


function WindowsManagementObject(const AWindowsManagementObject: IWindowsManagementObject): Variant;
begin
  VarClear(Result);

  { Assign the new variant the var type that was allocated for us }
  TWindowsManagementObjectVarData(Result).VType := FWindowsManagementVariantType.VarType;

  TWindowsManagementObjectVarData(Result).WmiObject := AWindowsManagementObject;
end;

//------------------------------------------------------------------------------
//------------------------------------------------------------------------------
//------------------------------------------------------------------------------
var
  OldVarDispProc: TVarDispProc;

procedure MyVarDispProc(Result: PVariant; const Instance: Variant; CallDesc: PCallDesc; Params: Pointer); cdecl;
var
  WMObject: IWindowsManagementObject;
  WmiObject: Variant;
begin
  if (TVarData(Instance).VType = varUnknown) and (Supports(IUnknown(TVarData(Instance).VUnknown), IWindowsManagementObject, WMObject)) then
  begin
    // Wrap interface in our custom variant type...
    WmiObject := WindowsManagementObject(WMObject);
    // ...and let the manager handle the method call
    FWindowsManagementVariantType.DispInvoke(@TVarData(Result^), TVarData(WmiObject), CallDesc, Params);
  end else
    OldVarDispProc(Result, Instance, CallDesc, Params);
end;

//------------------------------------------------------------------------------
//------------------------------------------------------------------------------
//------------------------------------------------------------------------------

initialization
  ComObj.CoInitFlags := COINIT_APARTMENTTHREADED;

  FWindowsManagementVariantType := TWindowsManagementObjectVariantType.Create;
  FWindowsManagementObjectMethodResultVariantType := TWindowsManagementObjectMethodResultVariantType.Create;

  OldVarDispProc := VarDispProc;
  VarDispProc := MyVarDispProc;

finalization
  FreeAndNil(FWindowsManagementVariantType);
  FreeAndNil(FWindowsManagementObjectMethodResultVariantType);

  VarDispProc := OldVarDispProc;
end.
