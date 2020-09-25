using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelGenerator.Models;
using System;
using System.Collections.Generic;

namespace ExcelGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            //teste();
            var t = new Class1();
            t.create(@"C:\Code\teste.xlsx");
            //t.CreateSheets(@"C:\Code\teste2.xlsx", "teste", 1);
            //t.CreateSheets(@"C:\Code\teste2.xlsx", "teste2", 2);
            //t.CreateSheets(@"C:\Code\teste2.xlsx", "teste3", 3);
            //t.ReadExcelFileSAX(@"C:\Code\teste.xlsx");
            t.WriteRandomValuesSAX(@"C:\Code\teste.xlsx",20,10,1,11);
           // t.WriteRandomValuesSAX(@"C:\Code\teste2.xlsx", 20, 10);
           // t.WriteRandomValuesSAX(@"C:\Code\teste.xlsx",10,10,2);
            //string path = @"C:\Code\teste.xlsx";
            //var docs = new DocumentCreator(path);

            //docs.CreateDocument();

            //docs.AddSheet("Teste");
            //docs.AddSheet("Teste2");
        }
        static void teste()
        {
            // Create a spreadsheet document by supplying the file name.  
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
                Create(@"C:\Code\teste.xlsx", SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.  
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.  
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.  
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.  
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "mySheet"
            };
            sheets.Append(sheet);

            // Close the document.  
            spreadsheetDocument.Close();

            Console.WriteLine("The spreadsheet document has been created.\nPress a key.");
            Console.ReadKey();

            TestModelList tmList = new TestModelList();
            tmList.TestData = new List<TestModel>();
            TestModel tm = new TestModel();
            tm.TestId = 1;
            tm.TestName = "Test1";
            tm.TestDesc = "Tested 1 time";
            tm.TestDate = DateTime.Now.Date;
            tmList.TestData.Add(tm);

            TestModel tm1 = new TestModel();
            tm1.TestId = 2;
            tm1.TestName = "Test2";
            tm1.TestDesc = "Tested 2 times";
            tm1.TestDate = DateTime.Now.AddDays(-1);
            tmList.TestData.Add(tm1);

            TestModel tm2 = new TestModel();
            tm2.TestId = 3;
            tm2.TestName = "Test3";
            tm2.TestDesc = "Tested 3 times";
            tm2.TestDate = DateTime.Now.AddDays(-2);
            tmList.TestData.Add(tm2);

            TestModel tm3 = new TestModel();
            tm3.TestId = 4;
            tm3.TestName = "Test4";
            tm3.TestDesc = "Tested 4 times";
            tm3.TestDate = DateTime.Now.AddDays(-3);
            tmList.TestData.Add(tm);

            var p = new CreateExcel();
            p.CreateExcelFile(tmList, @"C:\Code");
        }
    }
}
