using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelGenerator
{
    public class Class1
    {
        public void ReadExcelFileSAX(string fileName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
                string text;
                while (reader.Read())
                {
                    if (reader.ElementType == typeof(CellValue))
                    {
                        text = reader.GetText();
                        Console.Write(text + " ");
                    }
                }
                Console.WriteLine();
                Console.ReadKey();
            }
        }
        public void CreateSheets(string filename, string sheetname,int sheetid)
        {
            using (SpreadsheetDocument myDoc = SpreadsheetDocument.Open(filename, true))
            {
                WorkbookPart workbookPart = myDoc.WorkbookPart;
                if(sheetid == 1)
                {
                    workbookPart.Workbook = new Workbook();
                }
                Workbook workbook = myDoc.WorkbookPart.Workbook;
                //WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                OpenXmlWriter writer = OpenXmlWriter.Create(workbookPart);


                writer.WriteStartElement(workbook);
                writer.WriteStartElement(new Sheets());

                writer.WriteElement(new Sheet()
                {
                    Name = sheetname,
                    SheetId = (uint)sheetid
                });

                // this is for Sheets
                writer.WriteEndElement();
                // this is for Workbook
                writer.WriteEndElement();


                writer.Close();

                myDoc.Close();

            }
        }
        public void create(string fileName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart;
                OpenXmlWriter writer;
                OpenXmlWriter writer2;

                workbookPart = document.AddWorkbookPart();

                WorkbookStylesPart wbsp = workbookPart.AddNewPart<WorkbookStylesPart>();

                SharedStringTablePart stringTable = workbookPart.AddNewPart<SharedStringTablePart>();


                var workbook = new Workbook();
                workbookPart.Workbook = workbook;

                var fileVersion = new FileVersion() { ApplicationName = "Microsoft Office Excel" };
                workbook.Append(fileVersion);

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var relationshipId = workbookPart.GetIdOfPart(worksheetPart);
                var sheetPayment = new Sheet { Name = "Payment Upload", SheetId = 1, Id = relationshipId };

                var sheets = new Sheets();
                sheets.Append(sheetPayment);

                writer = OpenXmlWriter.Create(worksheetPart);

                writer.WriteStartElement(new Worksheet());
                writer.WriteStartElement(new SheetData());
                var worksheetPart2 = workbookPart.AddNewPart<WorksheetPart>();
                var relationshipId2 = workbookPart.GetIdOfPart(worksheetPart2);
                var sheetSuspense = new Sheet { Name = "Suspense Upload", SheetId = 2, Id = relationshipId2 };

                sheets.Append(sheetSuspense);

                writer2 = OpenXmlWriter.Create(worksheetPart2);

                writer2.WriteStartElement(new Worksheet());
                writer2.WriteStartElement(new SheetData());

                workbook.Append(sheets);

                writer.WriteEndElement();      // SheetData
                writer2.WriteEndElement();

                writer.WriteEndElement();      // Worksheet
                writer2.WriteEndElement();

                writer.Close();
                writer2.Close();

                document.WorkbookPart.Workbook.Save();
                document.Close();
            }
        }
        public void WriteRandomValuesSAX(string filename, int numRows, int numCols)
        {
            using (SpreadsheetDocument myDoc = SpreadsheetDocument.Open(filename, true))
            {
                WorkbookPart workbookPart = myDoc.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                string origninalSheetId = workbookPart.GetIdOfPart(worksheetPart);

                WorksheetPart replacementPart =
                workbookPart.AddNewPart<WorksheetPart>();
                string replacementPartId = workbookPart.GetIdOfPart(replacementPart);

                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
                OpenXmlWriter writer = OpenXmlWriter.Create(replacementPart);

                Row r = new Row();
                Cell c = new Cell();
                CellFormula f = new CellFormula();
                f.CalculateCell = true;
                f.Text = "RAND()";
                c.Append(f);
                CellValue v = new CellValue();
                c.Append(v);

                while (reader.Read())
                {
                    if (reader.ElementType == typeof(SheetData))
                    {
                        if (reader.IsEndElement)
                            continue;
                        writer.WriteStartElement(new SheetData());

                        for (int row = 0; row < numRows; row++)
                        {
                            writer.WriteStartElement(r);
                            for (int col = 0; col < numCols; col++)
                            {
                                writer.WriteElement(c);
                            }
                            writer.WriteEndElement();
                        }

                        writer.WriteEndElement();
                    }
                    else
                    {
                        if (reader.IsStartElement)
                        {
                            writer.WriteStartElement(reader);
                        }
                        else if (reader.IsEndElement)
                        {
                            writer.WriteEndElement();
                        }
                    }
                }

                reader.Close();
                writer.Close();

                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>()
                .Where(s => s.Id.Value.Equals(origninalSheetId)).First();
                sheet.Id.Value = replacementPartId;
                workbookPart.DeletePart(worksheetPart);
            }
        }
        public void WriteRandomValuesSAX(string filename, int numRows, int numCols, int sheetid, int countRows)
        {
            using (SpreadsheetDocument myDoc = SpreadsheetDocument.Open(filename, true))
            {
                WorkbookPart workbookPart = myDoc.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.Last();

                OpenXmlWriter writer = OpenXmlWriter.Create(worksheetPart);

                Row r = new Row();
                Cell c = new Cell();
                CellValue v = new CellValue("Test");
                c.AppendChild(v);

                var s = new SheetData();



                writer.WriteStartElement(new Worksheet());
                writer.WriteStartElement(s);
                //writer.WriteStartElement(new Sheet());
                for (int row = countRows; row < numRows; row++)
                {
                    writer.WriteStartElement(r);
                    for (int col = 0; col < numCols; col++)
                    {
                        writer.WriteElement(c);
                    }
                    writer.WriteEndElement();
                }
                //writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndElement();

                writer.Close();

                //writer = OpenXmlWriter.Create(myDoc.WorkbookPart);
                //writer.WriteStartElement(new Workbook());
                //writer.WriteStartElement(new Sheets());

                //writer.WriteElement(new Sheet()
                //{
                //    Name = $"Sheet21{sheetid}",
                //    SheetId = (uint)sheetid,
                //    Id = myDoc.WorkbookPart.GetIdOfPart(worksheetPart)
                //});

                //// this is for Sheets
                //writer.WriteEndElement();
                //// this is for Workbook
                //writer.WriteEndElement();


                //writer.Close();

                myDoc.Close();
            }
        }
    }
}