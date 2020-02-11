using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelToObjects
{
    public class ExcellToObjects : IDisposable
    {
        private List<Head> headList;
        private SpreadsheetDocument spreadsheetDocument;

        public ExcellToObjects(string file)
        {
            headList = new List<Head>();
            spreadsheetDocument = SpreadsheetDocument.Open(file, false);
        }

        public List<T> ConvertToObject<T>()
        {
            IEnumerable<Row> allRows = GetRows(spreadsheetDocument);
            var headRow = allRows.ElementAt(0);
            var dataRows = allRows.Where(r => r.RowIndex > 1);
            SetHead(headRow);
            List<T> data = ConvertRowsIntoObjects<T>(dataRows);

            return data;
        }

        private List<T> ConvertRowsIntoObjects<T>(IEnumerable<Row> rows)
        {
            List<T> data = new List<T>();
            foreach (Row row in rows)
            {
                T item = GetItem<T>(row, rows.ElementAt(0), spreadsheetDocument);
                data.Add(item);
            }

            return data;
        }

        private void SetHead(Row headRow)
        {
            foreach (Cell cell in headRow)
            {
                headList.Add(GetCellHead(cell));
            }
        }

        private Head GetCellHead(Cell cell)
        {
            SharedStringTablePart stringTablePart = spreadsheetDocument.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue.InnerXml;
            string reference = cell.CellReference.InnerText;
            reference = new String(reference.Where(c => c != '-' && (c < '0' || c > '9')).ToArray());

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                value = stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            return new Head() { Name = value, ReferenceColl = reference };
        }

        private string GetCellValue(SpreadsheetDocument spreadsheetDocument, Cell cell)
        {
            SharedStringTablePart stringTablePart = spreadsheetDocument.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue.InnerXml;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            return value;
        }

        private T GetItem<T>(Row row, Row head, SpreadsheetDocument spreadSheetDocument)
        {
            Type temp = typeof(T);
            T objReturn = Activator.CreateInstance<T>();

            foreach (var item in row.Descendants<Cell>())
            {
                var input = item.CellReference.InnerText.OnlyLetters();
                var column = headList.FirstOrDefault(h => h.ReferenceColl == input);
                var property = temp.GetProperties().FirstOrDefault(p => p.Name == column.Name);
                SetValueForProperty(property, objReturn, item);
            }

            return objReturn;
        }

        private void SetValueForProperty<T>(PropertyInfo property, T objReturn, Cell item)
        {
            if (property.PropertyType == typeof(String))
                property.SetValue(objReturn, GetCellValue(spreadsheetDocument, item), null);
            
            if (property.PropertyType == typeof(int))
                property.SetValue(objReturn, Convert.ToInt32(GetCellValue(spreadsheetDocument, item)), null);
           
            if (property.PropertyType == typeof(decimal?))
                property.SetValue(objReturn, Convert.ToDecimal(GetCellValue(spreadsheetDocument, item)), null);
            
            if (property.PropertyType == typeof(decimal))
                property.SetValue(objReturn, Convert.ToDecimal(GetCellValue(spreadsheetDocument, item)), null);

            //Tratar outros tipos aqui   ;)
        }

        private IEnumerable<Row> GetRows(SpreadsheetDocument spreadsheetDocument)
        {
            IEnumerable<Sheet> sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(relationshipId);
            Worksheet workSheet = worksheetPart.Worksheet;
            SheetData sheetData = workSheet.GetFirstChild<SheetData>();
            IEnumerable<Row> rows = sheetData.Descendants<Row>();
            return rows;
        }

        public void Dispose()
        {
            headList = null;
            spreadsheetDocument = null;
        }
    }

    public class Head
    {
        public string Name { get; set; }
        public string ReferenceColl { get; set; }
    }

    public static class StringExtensions
    {
        public static string OnlyLetters(this string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            string cleaned = new Regex(@"[^A-Z]+").Replace(s, "");
            return cleaned;
        }
    }
}
