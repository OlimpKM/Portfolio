unit DatabaseManager;

interface

uses
  Windows, Messages, SysUtils, Classes, DB, DBClient, Controls, Dialogs,
  Variants, ZDbcIntfs, Forms, Math, StrUtils, DateUtils,
  ZSqlUpdate, ZAbstractDataset, ZDataset, ZConnection, ZSqlProcessor,
  ZScriptParser, ZAbstractRODataset, ZStoredProcedure, ZAbstractConnection;

type
  { Record representing a dynamic query and its metadata }
  TQueryRecord = record
    Query       : TZQuery;
    Update      : TZUpdateObject;
    IsPrepared  : Boolean;
    InternalID  : Integer;
    Description : String;
  end;
  PQueryRecord = ^TQueryRecord;

  { Record for batch SQL commands }
  TSqlCommand = record
    Executed : Boolean;
    Command  : String;
    Query    : TZQuery;
    Status   : Boolean;
    Duration : Int64;
    Notes    : String;
  end;
  PSqlCommand = ^TSqlCommand;

  { Main database management class using Zeos components }
  TDatabaseManager = class(TObject)
  private
    FConnection     : TZConnection;
    FFormatNumeric  : String;
    FFormatDate     : String;
    FFormatTime     : String;

    procedure _ApplyDisplayFormats(ARec: PQueryRecord);
    procedure _SetReadOnly(ARec: PQueryRecord; ReadOnly: Boolean);
  protected
    procedure HandleError(const MethodName, Msg: String); virtual;
  public
    constructor Create(AConnection: TZConnection);

    function ExecuteQuery(const SQL: String; var ARec: PQueryRecord): Boolean;
    function PrepareParameters(var ARec: PQueryRecord; const Params: array of Variant): Boolean;

    property Connection: TZConnection read FConnection;
    property FormatDate: String read FFormatDate write FFormatDate;
  end;

implementation

{ TDatabaseManager }

constructor TDatabaseManager.Create(AConnection: TZConnection);
begin
  inherited Create;
  FConnection    := AConnection;
  FFormatNumeric := '#,##0.00';
  FFormatDate    := 'yyyy-MM-dd';
  FFormatTime    := 'HH:mm:ss';
end;

function TDatabaseManager.ExecuteQuery(const SQL: String; var ARec: PQueryRecord): Boolean;
begin
  Result := False;
  try
    if not Assigned(FConnection) then Exit;

    if not Assigned(ARec^.Query) then
      ARec^.Query := TZQuery.Create(nil);

    ARec^.Query.Connection := FConnection;
    ARec^.Query.SQL.Text   := SQL;
    ARec^.Query.Open;

    _ApplyDisplayFormats(ARec);
    Result := True;
  except
    on E: Exception do
      HandleError('ExecuteQuery', E.Message);
  end;
end;

function TDatabaseManager.PrepareParameters(var ARec: PQueryRecord; const Params: array of Variant): Boolean;
var
  i: Integer;
begin
  Result := False;
  try
    if not Assigned(ARec^.Query) then Exit;

    for i := Low(Params) to High(Params) do
    begin
      // Logic for mapping variant array to query parameters
      if i < ARec^.Query.Params.Count then
        ARec^.Query.Params[i].Value := Params[i];
    end;

    ARec^.IsPrepared := True;
    Result := True;
  except
    on E: Exception do
      HandleError('PrepareParameters', E.Message);
  end;
end;

procedure TDatabaseManager._ApplyDisplayFormats(ARec: PQueryRecord);
var
  i: Integer;
  Field: TField;
begin
  { Iterate through fields and apply localized formatting based on data type }
  for i := 0 to ARec^.Query.FieldCount - 1 do
  begin
    Field := ARec^.Query.Fields[i];
    case Field.DataType of
      ftFloat, ftCurrency:
        (Field as TNumericField).DisplayFormat := FFormatNumeric;
      ftDate:
        (Field as TDateField).DisplayFormat := FFormatDate;
      ftTime:
        (Field as TTimeField).DisplayFormat := FFormatTime;
      ftDateTime:
        (Field as TDateTimeField).DisplayFormat := FFormatDate + ' ' + FFormatTime;
    end;
  end;
end;

procedure TDatabaseManager._SetReadOnly(ARec: PQueryRecord; ReadOnly: Boolean);
begin
  if Assigned(ARec^.Query) then
    ARec^.Query.ReadOnly := ReadOnly;
end;

procedure TDatabaseManager.HandleError(const MethodName, Msg: String);
begin
  // Generic error logging - to be overridden or implemented with a logger
  OutputDebugString(PChar(Format('[%s] Error: %s', [MethodName, Msg])));
end;

end.