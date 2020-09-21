using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelGenerator
{
    public class DocumentCreator
    {
        private SpreadsheetDocument Document { get; set; }
        private WorkbookPart Workbookpart { get; set; }

        private readonly string _path;

        public DocumentCreator(string path)
        {
            _path = path;
        }

        public void CreateDocument()
        {
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
                 Create(_path, SpreadsheetDocumentType.Workbook);

            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());


            spreadsheetDocument.Close();
        }
        public void CreateSheetData(IEnumerable<string> headers,string sheetName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(_path, true))
            {
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook
                    .GetFirstChild<Sheets>();

                var sheet = sheets.ChildElements.Where(x => x.LocalName == sheetName);

                SheetData sheetData = new SheetData();

                sheetData.Append(CreateHeaderRow(headers));
            }

        }

        private Row CreateHeaderRow(IEnumerable<string> headers)
        {
            Row rows = new Row();
            foreach (var header in headers)
            {
                rows.Append(CreateCell(header));

            }
            return rows;
        }
        private Cell CreateCell(string text)
        {
            Cell cell = new Cell
            {
                StyleIndex = 1U,
                CellValue = new CellValue(text)
            };
            return cell;
        }
        public void AddSheet(string sheetName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(_path, true))
            {
                // Add a blank WorksheetPart.  
                WorksheetPart newWorksheetPart =
                    spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook
                    .GetFirstChild<Sheets>();
                string relationshipId = spreadsheetDocument.WorkbookPart
                    .GetIdOfPart(newWorksheetPart);

                // Get a unique ID for the new worksheet.  
                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Count() > 0)
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;


                // Append the new worksheet and associate it with the workbook.  
                Sheet sheet = new Sheet()
                {
                    Id = relationshipId,
                    SheetId = sheetId,
                    Name = sheetName
                };
                sheets.Append(sheet);
            }


        }
    }
}
