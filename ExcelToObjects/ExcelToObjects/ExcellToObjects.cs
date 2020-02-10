using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToObjects
{
    public class ExcellToObjects
    {


        public static List<T> ConvertToObject<T>(string file) 
        {
            List<T> data = new List<T>();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(file, false))
            {
                IEnumerable<Row> rows = GetRows(spreadsheetDocument);
                var headList = GetHead(spreadsheetDocument, rows);

                foreach (Row row in rows)
                {
                    T item = GetItem<T>(row, rows.ElementAt(0));
                    data.Add(item);
                }
            }

            return data;

        }

        private static List<Head> GetHead(SpreadsheetDocument spreadsheetDocument, IEnumerable<Row> rows)
        {
            List<Head> heads = new List<Head>();
            foreach (Cell cell in rows.ElementAt(0))
            {
                heads.Add(GetCellHead(spreadsheetDocument, cell));

            }
            return heads;
        }

        private static Head GetCellHead(SpreadsheetDocument spreadsheetDocument, Cell cell)
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

        private static string GetCellValue(SpreadsheetDocument spreadsheetDocument, Cell cell)
        {
            SharedStringTablePart stringTablePart = spreadsheetDocument.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue.InnerXml;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            return value;
        }

        private static T GetItem<T>(Row row, Row head)
        {
            throw new NotImplementedException();
        }

        private static IEnumerable<Row> GetRows(SpreadsheetDocument spreadsheetDocument)
        {
            IEnumerable<Sheet> sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(relationshipId);
            Worksheet workSheet = worksheetPart.Worksheet;
            SheetData sheetData = workSheet.GetFirstChild<SheetData>();
            IEnumerable<Row> rows = sheetData.Descendants<Row>();
            return rows;
        }
    }

    public class Head
    {
        public string Name { get; set; }
        public string ReferenceColl { get; set; }
    }

}
