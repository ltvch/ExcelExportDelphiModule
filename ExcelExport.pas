unit ExcelExport;

(* Export data to Excel as .xls or .csv format *)
(* use any component content DataSet *)
(* Version 0.1 stable *)
(* Develop by ltvch 2017 *)

interface

uses System.SysUtils, System.Classes, System.Variants, Data.DB,
  Vcl.Forms, Vcl.Dialogs, Vcl.Controls, ComObj, ActiveX, Excel2000,
  Vcl.ComCtrls, Vcl.ExtCtrls,  Vcl.StdCtrls, DateUtils, Graphics;

procedure SaveAsDialog(const TypeFormatExport: TRadioGroup; ThemeValue: string;
  Data: TDataSet); overload;
procedure SaveAsDialog(const ThemeValue: string; Data: TDataSet); overload;

type
  TExcelPrepare = class
  strict private
    fDataSet: TDataSet;
    fSheetName: string;
    function IsOLEObjectInstalled(Name: String): boolean;
    function GetColumn(const Col: integer): String;
    function GetPrepareRange(RowFirst, ColFirst, RowEnd,
      ColEnd: integer): String;
    function AddToString(const ABaseValue, AFieldValue: string): string;

    procedure DecorateTotalSection(const Sheet, StartCell: Variant;
      Total: string);
    procedure BuildTitle(const Sheet: Variant; fDataSet: TDataSet);
    procedure BuildTotal(const Sheet: Variant; fDataSet: TDataSet);
    procedure FillingExcel(const Sheet: Variant; fDataSet: TDataSet);
  public
    procedure SaveToCSVFile(const AFileName: TFileName);
    procedure SaveToExcelFile(const AFileName: TFileName);
    constructor Create(Sheet: string; Data: TDataSet); overload;
  end;

implementation

{ uses uLog; }

procedure SaveAsDialog(const ThemeValue: string; Data: TDataSet);
{ true and work }
var
  ExportObj: TExcelPrepare;
  SaveDialog: TSaveDialog;
begin
  (* -create ExportObj where set Sheet name and put Data - *)
  (* use so DBGridMain.DataSource.DataSet or DataSource.DataSet as DATA *)
  ExportObj := TExcelPrepare.Create(ThemeValue, Data);
  (* create SaveDialog for saving anywhere *)
  SaveDialog := TSaveDialog.Create(nil);
  (* filename suggestion *)
  SaveDialog.FileName := ThemeValue + ' созданный ' + DateToStr(Date()) + '_' +
    inttostr(SecondOf(now())) + '.xls';
  (* by how items are selected in radiogroup *)
  try
    begin
      SaveDialog.Filter := 'Excel File|*.xls';
      (* if SaveDialog gives us a way for for save *)
      if SaveDialog.Execute then
        ExportObj.SaveToExcelFile(SaveDialog.FileName)
    end;
  finally
    (* remove objects to clean memory *)
    FreeAndNil(SaveDialog);
    FreeAndNil(ExportObj);
  end;
end;

procedure SaveAsDialog(const TypeFormatExport: TRadioGroup; ThemeValue: string;
  Data: TDataSet);
{ true and work }
var
  ExportObj: TExcelPrepare;
  SaveDialog: TSaveDialog;
begin
  (* -create ExportObj, set Sheet name and put Data - *)
  (* use so Query.DataSet or DBGridMain.DataSource.DataSet or DataSource.DataSet as DATA *)
  ExportObj := TExcelPrepare.Create(ThemeValue, Data);
  (* create SaveDialog for saving anywhere *)
  SaveDialog := TSaveDialog.Create(nil);
  (* filename suggestion *)
  SaveDialog.FileName := ThemeValue + ' созданный ' + DateToStr(Date()) + '_' +
    inttostr(SecondOf(now()));
  (* by how items are selected in radiogroup *)
  try
    case TypeFormatExport.ItemIndex of
      (* so saves the file like selected type *)
      0:
        begin
          SaveDialog.Filter := 'Excel File|*.xls';
          (* if SaveDialog gives us a way for for save *)
          if SaveDialog.Execute then
            ExportObj.SaveToExcelFile(SaveDialog.FileName)
        end;
      1:
        begin
          SaveDialog.Filter := 'Comma Delimited|*.csv';
          if SaveDialog.Execute then
            ExportObj.SaveToCSVFile(SaveDialog.FileName);
        end;
    end;
  finally
    (* remove objects to clean memory *)
    FreeAndNil(SaveDialog);
    FreeAndNil(ExportObj);
  end;
end;

{ TExcelExport }

constructor TExcelPrepare.Create(Sheet: string; Data: TDataSet);
begin
  if not IsOLEObjectInstalled('Excel.Application') then
  begin
    MessageDlg('MS Excel not installed yet. Can`t work.', mtERROR, [mbok], 0);
    Exit;
  end;

  if Assigned(Data) then
    fDataSet := Data
  else
    raise Exception.Create('Not find data in constructor!');

  if (Sheet <> '') then
    fSheetName := Sheet
  else
    raise Exception.Create('Not find Sheet in constructor!');
end;

function TExcelPrepare.IsOLEObjectInstalled(Name: String): boolean;
var
  ClassID: TCLSID;
begin
  Result := CLSIDFromProgID(PWideChar(WideString(name)), ClassID) = S_OK;
end;

function TExcelPrepare.AddToString(const ABaseValue,
  AFieldValue: string): string;
begin
  if ABaseValue = '' then
    Result := QuotedStr(AFieldValue)
  else
    Result := ABaseValue + ',' + QuotedStr(AFieldValue);
end;

procedure TExcelPrepare.SaveToCSVFile(const AFileName: TFileName);
var
  RowValue: string;
  DataCols, I: integer;
  Stream: TMemoryStream;
begin
  DataCols := fDataSet.FieldCount;

  Stream := TMemoryStream.Create;
  try
    (* write the titles *)
    for I := 0 to DataCols - 1 do
      RowValue := AddToString(RowValue, fDataSet.Fields[I].FieldName);
    RowValue := RowValue + #13#10;
    Stream.Write(Pointer(RowValue)^, Length(RowValue) * SizeOf(Char));

    (* write data *)
    fDataSet.DisableControls;
    fDataSet.First;

    while not fDataSet.Eof do
    begin
      RowValue := '';
      for I := 0 to DataCols - 1 do
        RowValue := AddToString(RowValue, fDataSet.Fields[I].AsString);
      RowValue := RowValue + #13#10;
      Stream.Write(Pointer(RowValue)^, Length(RowValue) * SizeOf(Char));

      fDataSet.Next;
    end;
    Stream.SaveToFile(AFileName);
  finally
    FreeAndNil(Stream);
  end;
end;

procedure TExcelPrepare.SaveToExcelFile(const AFileName: TFileName);
var
  ExcelApp, Sheet: Variant;
  NameSheet: string;
begin
  NameSheet := Copy(fSheetName, 0, 31); // WorkSheets name must bee <= 31
  (* desable cotrols  => dataset for get access to stable value *)
  fDataSet.DisableControls;
  (* move to last row in DataSet for get All rows *)
  fDataSet.Last;

  try
    (* create Excel table *)
    ExcelApp := CreateOleObject('Excel.Application');
    ExcelApp.Visible := False;
    // ExcelApp.WorkBooks.Add(-4167); // implement WorkBook
    ExcelApp.WorkBooks.Add; // so work without magic digits
    (* change all RowHeight for fixed field with long valuest *)
    ExcelApp.WorkBooks[1].WorkSheets[1].Columns['A:Z'].RowHeight := 42;
    (* create sheet, set its name in Excel table and implement sheet variable *)
    ExcelApp.WorkBooks[1].WorkSheets[1].Name := NameSheet;
    Sheet := ExcelApp.WorkBooks[1].WorkSheets[NameSheet];

    BuildTitle(Sheet, fDataSet);
    BuildTotal(Sheet, fDataSet);

    FillingExcel(Sheet, fDataSet);

    try
      (* Save Excel file as FileName *)
      ExcelApp.WorkBooks[1].SaveAs(AFileName);
    except
      on E: Exception do
        raise Exception.Create('Data transfer error: ' + E.Message);
      // sLog('', E.ClassName + ' ' + E.Message);
    end;

    (* enable DataSet for use higest *)
    fDataSet.EnableControls;

  finally
    if not VarIsEmpty(ExcelApp) then
    begin
      (* Make visible MS Excel *)
      ExcelApp.Visible := True;
      (* Unlink our variable from the application. *)
      ExcelApp := UnAssigned;
      (* Open Workbook in MS Excel -> not need on this iteration now *)
      // ExcelApp.WorkBooks.Open(AFileName);
      // ExcelApp.WorkBooks.Close; //Close work book.
      // ExcelApp.Quit;//Close MS Excel
    end;
  end;
end;

(* Auxiliary - generates an alphabetical representation *)
(* (Excel column like A or AJ name) of a column by its number *)
function TExcelPrepare.GetColumn(const Col: integer): String;
begin
  Result := '';
  if Col < 27 then
    Result := CHR(Col + 64)
  else
    Result := CHR((Col div 26) + 64) + CHR((Col mod 26) + 64);
end;

(* Returns a string of type like 'A1: C5' for the range *)
(* (RowFirst, ColFirst - TopLeft corner, RowEnd, ColEnd - RightBottom corner) *)
function TExcelPrepare.GetPrepareRange(RowFirst, ColFirst, RowEnd,
  ColEnd: integer): String;
begin
  Result := GetColumn(ColFirst) + inttostr(RowFirst) + ':' + GetColumn(ColEnd) +
    inttostr(RowEnd);
end;

procedure TExcelPrepare.BuildTotal(const Sheet: Variant; fDataSet: TDataSet);
var
  Count, Indent: integer;
  I, J: integer;
  Cell: string;
begin
  Indent := 2;
  Count := fDataSet.RecordCount;

  DecorateTotalSection(Sheet, Count + Indent, 'Общие данные');
  DecorateTotalSection(Sheet, Count + Indent + 7, 'Абсолютные значения');
  DecorateTotalSection(Sheet, Count + Indent + 14,
    'Абсолютные значения без нулевых значений');

  for I := Indent to fDataSet.FieldCount do
  begin

    for J := 0 to Count do
      Cell := GetPrepareRange(I, I, Count + 1, I);

    Sheet.Cells[Count + 4, I].FormulaLocal := Format('=СЧЁТ(%s)', [Cell]);
    Sheet.Cells[Count + 5, I].FormulaLocal := Format('=СУММ(%s)', [Cell]);
    Sheet.Cells[Count + 6, I].FormulaLocal := Format('=СРЗНАЧ(%s)', [Cell]);
    Sheet.Cells[Count + 7, I].FormulaLocal := Format('=МАКС(%s)', [Cell]);
    Sheet.Cells[Count + 8, I].FormulaLocal := Format('=МИН(%s)', [Cell]);

    Sheet.Cells[Count + 11, I].FormulaLocal := Format('=ABS(СЧЁТ(%s))', [Cell]);
    Sheet.Cells[Count + 12, I].FormulaLocal := Format('=ABS(СУММ(%s))', [Cell]);
    Sheet.Cells[Count + 13, I].FormulaLocal :=
      Format('=ABS(СРЗНАЧ(%s))', [Cell]);
    Sheet.Cells[Count + 14, I].FormulaLocal := Format('=ABS(МАКС(%s))', [Cell]);
    Sheet.Cells[Count + 15, I].FormulaLocal := Format('=ABS(МИН(%s))', [Cell]);

    Sheet.Cells[Count + 18, I].FormulaLocal :=
      Format('=ABS(СЧЁТЕСЛИ(%s;">0"))', [Cell]);
    Sheet.Cells[Count + 19, I].FormulaLocal :=
      Format('=ABS(СУММЕСЛИ(%s;">0"))', [Cell]);
    Sheet.Cells[Count + 20, I].FormulaLocal :=
      Format('=ABS(СРЗНАЧ(%s))', [Cell]);
    Sheet.Cells[Count + 21, I].FormulaLocal := Format('=ABS(МАКС(%s))', [Cell]);
    Sheet.Cells[Count + 22, I].FormulaLocal := Format('=ABS(МИН(%s))', [Cell]);
  end;
end;

procedure TExcelPrepare.BuildTitle(const Sheet: Variant; fDataSet: TDataSet);
var
  Index: byte;
  I: integer;
begin
  (* make first row as title of report *)
  Index := 1; // use first row as title columns in Excel form
  (* set parameter in Title use row cells *)
  for I := 1 to fDataSet.FieldCount do
  begin
    Sheet.Cells[Index, I] := fDataSet.Fields[I - 1].FieldName;
    // add value use item
    Sheet.Cells[Index, I].Font.Bold := True;
    // Sheet.Cells[Index, I].Font.Color := clNavy;
    // sheet.Columns[I].ColumnWidth :=
    // DBGridMain.Columns.Items[I-1].Field.DisplayWidth+1;
    Sheet.Cells[Index, I].EntireColumn.AutoFit;
    Sheet.Columns[I].WrapText := True;
    // Sheet.Columns[I].ColumnWidth := 16;
    // Sheet.Columns[I].RowHeight := Sheet.Cells[1, I].Font.Size + 2;
    Sheet.Columns[I].HorizontalAlignment := -4108;
    Sheet.Columns[I].VerticalAlignment := -4108;
  end;
end;

procedure TExcelPrepare.FillingExcel(const Sheet: Variant; fDataSet: TDataSet);
var
  Index, I, J: integer;
begin
  (* from first row => add value in column Excel *)
  Index := 2;
  fDataSet.First;
  for I := 1 to fDataSet.RecordCount do
  begin
    for J := 1 to fDataSet.FieldCount do
      if (fDataSet.Fields[J - 1].DataType in [ftDate]) then
      begin
        // Sheet.Cells[Index, J].NumberFormat := 'ЧЧ:ММ:СС';
        Sheet.Cells[Index, J] := fDataSet.Fields[J - 1].AsDateTime;
      end
      else if (fDataSet.Fields[J - 1].DataType in [ftShortint, ftSmallint,
        ftInteger, ftLargeint, ftWord]) then
        Sheet.Cells[Index, J] := fDataSet.Fields[J - 1].AsInteger

      else if (fDataSet.Fields[J - 1].DataType in [ftSingle, ftFloat,
        ftCurrency, ftVariant, ftBCD]) then
        (* if value in grid and his type is floas = '' so default in excel set 0. but we wants '' *)
        if (fDataSet.Fields[J - 1].AsFloat <> 0) then
        begin
          Sheet.Cells[Index, J] := fDataSet.Fields[J - 1].AsFloat
        end
        else
        begin
          Sheet.Cells[Index, J].NumberFormatLocal := '@';
          Sheet.Cells[Index, J] := AnsiString(' ');
        end
      else
        Sheet.Cells[Index, J] := fDataSet.Fields[J - 1].AsString;

    Inc(Index);
    fDataSet.Next;
  end;
end;

procedure TExcelPrepare.DecorateTotalSection(const Sheet, StartCell: Variant;
  Total: string);
var
  I: integer;
const
  Row: array [0 .. 4] of string = ('Кол-во', 'Сумма', 'Среднее', 'Максимум',
    'Минимум');
begin
  Sheet.Cells[StartCell, 1].Value := ''; // total
  Sheet.Cells[StartCell + 1, 1].Value := Total; // title bottom total
  Sheet.Cells[StartCell + 1, 1].Font.Bold := True;
  Sheet.Range[Sheet.Cells.Item[StartCell + 1, 1],
    Sheet.Cells.Item[StartCell + 1, 2]].MergeCells := True;

  for I := 0 to Length(Row) - 1 do
  begin
    Sheet.Cells[StartCell + 2 + I, 1].Value := Row[I];
    // indent after total and title
    Sheet.Cells[StartCell + 2 + I, 1].Font.Bold := True;
  end;
end;

end.
