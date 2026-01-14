using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace WinCo.Core.Services
{
    /// <summary>
    /// Nowoczesna us³uga eksportu danych do formatu Excel (XLSX)
    /// z wykorzystaniem biblioteki OpenXML.
    /// Oparte na rozwi¹zaniu Mike's Knowledge Base, zrefaktoryzowane pod k¹tem nowoczesnych standardów C# i wydajnoœci.
    /// </summary>
    public class ExcelExportService
    {
        /// <summary>
        /// Generuje dokument Excel na podstawie generycznej listy obiektów.
        /// </summary>
        public bool ExportListToExcel<T>(List<T> data, string filePath, string sheetName = "Raport")
        {
            try
            {
                var dataTable = ConvertListToDataTable(data);
                dataTable.TableName = sheetName;

                using (var document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
                {
                    BuildExcelStructure(dataTable, document);
                }
                return true;
            }
            catch (Exception ex)
            {
                // W prawdziwym systemie warto tu dodaæ Logger (np. NLog lub Serilog)
                System.Diagnostics.Trace.WriteLine($"Excel Export Error: {ex.Message}");
                return false;
            }
        }

        private void BuildExcelStructure(DataTable dt, SpreadsheetDocument document)
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            var workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            workbookStylesPart.Stylesheet = new Stylesheet(); // Podstawowy styl wymagany przez Excel

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            var sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
            var sheet = new Sheet()
            {
                Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = dt.TableName ?? "Sheet1"
            };
            sheets.Append(sheet);

            PopulateSheetData(dt, worksheetPart);
            workbookPart.Workbook.Save();
        }

        private void PopulateSheetData(DataTable dt, WorksheetPart worksheetPart)
        {
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            // 1. Tworzenie nag³ówków
            var headerRow = new Row();
            foreach (DataColumn column in dt.Columns)
            {
                headerRow.Append(CreateCell(column.ColumnName, CellValues.String));
            }
            sheetData.Append(headerRow);

            // 2. Dodawanie danych
            foreach (DataRow row in dt.Rows)
            {
                var newRow = new Row();
                foreach (var item in row.ItemArray)
                {
                    newRow.Append(CreateCell(item?.ToString() ?? string.Empty));
                }
                sheetData.Append(newRow);
            }
        }

        private Cell CreateCell(string value, CellValues dataType = CellValues.String)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = dataType
            };
        }

        private DataTable ConvertListToDataTable<T>(List<T> items)
        {
            var dt = new DataTable(typeof(T).Name);
            var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            foreach (var prop in properties)
            {
                var propType = prop.PropertyType;
                if (propType.IsGenericType && propType.GetGenericTypeDefinition() == typeof(Nullable<>))
                    propType = Nullable.GetUnderlyingType(propType);

                dt.Columns.Add(prop.Name, propType);
            }

            foreach (var item in items)
            {
                var values = properties.Select(p => p.GetValue(item, null) ?? DBNull.Value).ToArray();
                dt.Rows.Add(values);
            }

            return dt;
        }
    }
}
