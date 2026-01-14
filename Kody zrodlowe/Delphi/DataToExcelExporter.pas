unit DataToExcelExporter;

interface

uses
  Windows, Messages, SysUtils, Classes, ActiveX, Forms,
  Dialogs, ADODB, Db, ADOX_TLB, DBGrids, Math;

type
  TProgressEvent = procedure(Sender : TObject; Percent : Integer) of object;

  { Container to preserve dataset state during export }
  TDataSetState = class
    BeforeScrollEvent      : TDataSetNotifyEvent;
    AfterScrollEvent       : TDataSetNotifyEvent;
    AutoCalcFieldsProperty : Boolean;
  end;

  { Main class for exporting data to MS Excel via ADO }
  TExcelExporter = class
  private
    FLocked          : Boolean;
    FShowDialog      : Boolean;
    FInitialDir      : String;
    FFileName        : String;
    FSheetName       : String;
    FOnProgress      : TProgressEvent;

    FADOConnection   : TADOConnection;
    FADODataSet      : TADODataSet;

    function _DbGridToExcel(Grid : TDBGrid; FileName, SheetName: String; UseDataTypeTag : Boolean) : Boolean;
    function _DataSetToExcel(Data: TDataSet; FileName, SheetName: String; UseDataTypeTag : Boolean = False) : Boolean;
    function _GetExcelConnectionString(FileName: String; ReadOnly: Boolean = False): String;
    function _GetDataTypeTag(Field: TField): String;
  public
    constructor Create;
    destructor Destroy; override;

    function ExportGrid(Grid: TDBGrid; const DestFileName: String; const SheetName: String = 'Data'): Boolean;

    property OnProgress: TProgressEvent read FOnProgress write FOnProgress;
    property InitialDir: String read FInitialDir write FInitialDir;
  end;

implementation

{ TExcelExporter }

constructor TExcelExporter.Create;
begin
  FShowDialog := True;
  FLocked := False;
end;

destructor TExcelExporter.Destroy;
begin
  inherited;
end;

function TExcelExporter._GetExcelConnectionString(FileName: String; ReadOnly: Boolean): String;
begin
  { Standard connection string for Excel via ACE/Jet OLEDB }
  Result := 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + FileName +
            ';Extended Properties="Excel 12.0 Xml;HDR=YES";';
end;

function TExcelExporter._DbGridToExcel(Grid: TDBGrid; FileName, SheetName: String; UseDataTypeTag: Boolean): Boolean;
var
  i, j, RowCount: Integer;
  ColumnList, CreateTableSQL: String;
  ADOQuery: TADOQuery;
  Catalog: _Catalog;
  SavePlace: TBookmark;
begin
  Result := True;
  if not Assigned(Grid.DataSource) or not Assigned(Grid.DataSource.DataSet) then Exit(False);

  { Use ADOX to create the Excel file structure }
  Catalog := CoCatalog.Create;
  try
    try
      Catalog.Create(_GetExcelConnectionString(FileName));
    except
      Exit(False);
    end;

    ADOQuery := TADOQuery.Create(nil);
    try
      ADOQuery.ConnectionString := _GetExcelConnectionString(FileName);

      { Build Dynamic CREATE TABLE query based on Grid columns }
      ColumnList := '';
      for i := 0; i < Grid.Columns.Count do
      begin
        if ColumnList <> '' then ColumnList := ColumnList + ', ';
        ColumnList := ColumnList + '[' + Grid.Columns[i].Title.Caption + '] MEMO';
      end;

      CreateTableSQL := 'CREATE TABLE [' + SheetName + '] (' + ColumnList + ')';
      ADOQuery.SQL.Text := CreateTableSQL;
      ADOQuery.ExecSQL;

      { Data Transfer }
      ADOQuery.SQL.Text := 'SELECT * FROM [' + SheetName + ']';
      ADOQuery.Open;

      SavePlace := Grid.DataSource.DataSet.GetBookmark;
      Grid.DataSource.DataSet.DisableControls;
      try
        Grid.DataSource.DataSet.First;
        RowCount := Grid.DataSource.DataSet.RecordCount;

        while not Grid.DataSource.DataSet.Eof do
        begin
          ADOQuery.Append;
          for i := 0; i < Grid.Columns.Count do
          begin
            { Mapping logic between DataSet and Excel cell }
            if Assigned(Grid.Columns[i].Field) then
              ADOQuery.Fields[i].Value := Grid.Columns[i].Field.Value;
          end;
          ADOQuery.Post;

          { Update progress }
          if Assigned(FOnProgress) then
            FOnProgress(Self, Round((Grid.DataSource.DataSet.RecNo / RowCount) * 100));

          Grid.DataSource.DataSet.Next;
        end;
      finally
        Grid.DataSource.DataSet.GotoBookmark(SavePlace);
        Grid.DataSource.DataSet.FreeBookmark(SavePlace);
        Grid.DataSource.DataSet.EnableControls;
      end;

    finally
      ADOQuery.Close;
      ADOQuery.Free;
    end;
  finally
    Catalog := nil;
  end;
end;

function TExcelExporter.ExportGrid(Grid: TDBGrid; const DestFileName: String; const SheetName: String): Boolean;
begin
  Result := _DbGridToExcel(Grid, DestFileName, SheetName, False);
end;

function TExcelExporter._GetDataTypeTag(Field: TField): String;
begin
  { Logic to determine data type for formatting }
  Result := 'S'; // Default: String
  if Field is TNumericField then Result := 'N'
  else if Field is TDateTimeField then Result := 'D'
  else if Field is TBooleanField then Result := 'B';
end;

end.