using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Tesseract;

namespace csharp_dotnetcore_examples {
    class Program {
        static void Main (string[] args) {
            System.Console.WriteLine("teste");
            //ReadExcel("teste1.xlsx");
            ReadExcelClosedXml("teste_2.xlsx");
            // TesseractTest();

            // string path = "./tessdata/por.traineddata";
            // string pathOut = "por.traineddata.txt";
            // using(System.IO.FileStream sr = new System.IO.FileStream(path, System.IO.FileMode.Open)) {
            //     byte[] bs = new byte[sr.Length];
                
            //     sr.Read(bs, 0, bs.Length);
            //     string content = "";
            //     content = System.Convert.ToBase64String(bs);

            //     System.IO.File.WriteAllText(pathOut,content, System.Text.Encoding.UTF8);

            // }
        }
        public static void ReadExcelClosedXml(string path) {
            
            var wb = new XLWorkbook(path);
            
            var ws = wb.Worksheet(1);
            for (int i = 1; i < ws.RowsUsed().Count(); i++) {
                IXLCell c = ws.Cell(i, 1);
                var str = c.Value.ToString();
                System.Console.WriteLine("{0}", str);
            }
            
            wb.Dispose();
        }

        public static void ReadExcel (string path) {

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open (path, true)) {
                Sheets sheets = doc.WorkbookPart.Workbook.Sheets;
                WorkbookPart wPart = doc.WorkbookPart;

                IEnumerable<Sheet> shs = sheets.ChildElements.Cast<Sheet> ();
                var sheet1 = shs.FirstOrDefault<Sheet> (/*s => s.Name == "Planilha1"*/); // get tab by name

                Worksheet workSheet = ((WorksheetPart) wPart.GetPartById (sheet1.Id)).Worksheet;
                List<SheetData> rows = workSheet.ChildElements.OfType<SheetData> ().ToList ();

                string currCellValue = null;

                List<List<string>> lstSheet = new List<List<string>> (rows[0].ChildElements.Count);

                for (int i = 0; i < rows[0].ChildElements.Count; i++) {
                    lstSheet.Add (new List<string> (4));

                    Row currentrow = (Row) rows[0].ChildElements.GetItem (i);

                    Cell[] cells = new Cell[] {
                        (Cell) currentrow.ChildElements.GetItem (0)
                        //,(Cell) currentrow.ChildElements.GetItem (1)
                        //(Cell) currentrow.ChildElements.GetItem (2)
                        //(Cell) currentrow.ChildElements.GetItem (3)
                    };

                    foreach (Cell c in cells) {
                        currCellValue = getStringFromCellValue (wPart, c);
                        lstSheet.Last ().Add (currCellValue);
                    }
                }
                System.IO.StreamWriter strW = new System.IO.StreamWriter ("test_output.csv");
                foreach (var rs in lstSheet) {
                    foreach (var c in rs) {
                        Console.Write ("{0}; ", c);
                        strW.Write ("{0};", c);
                    }
                    strW.WriteLine ();
                    Console.WriteLine ();
                }
                strW.Dispose ();
                strW.Close ();
            }
#region  ............................ Inner functions ...............................
            
            string getStringFromCellValue (WorkbookPart wPart, Cell c) {
                string currCellValue = null;
                if (c.DataType != null) {
                    Console.WriteLine ("DataType: {0}", c.DataType.InnerText);
                    if (c.DataType == CellValues.SharedString) {
                        int id = -1;
                        if (int.TryParse (c.InnerText, out id)) {
                            SharedStringItem item = GetSharedStringItemById (wPart, id);
                            if (item.Text != null) {
                                currCellValue = item.Text.Text;
                            } else if (item.InnerText != null) {
                                currCellValue = item.InnerText;
                            } else if (item.InnerXml != null) {
                                currCellValue = item.InnerXml;
                            }
                        }
                    }
                } else {
                    Console.WriteLine ("DataType: {0}", c.DataType?.InnerText);
                    currCellValue = c.InnerText;
                }

                return currCellValue;
            }
            SharedStringItem GetSharedStringItemById (WorkbookPart workbookPart, int id) {
                return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem> ().ElementAt (id);
            }

#endregion ..........................................................................
        
        }
        public static void WriteExcel () {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Create (string.Format ("new_doc_{0:yyyy-MM-dd_HHmmss}.xlsx", DateTime.Now), SpreadsheetDocumentType.Workbook)) {
                WorkbookPart wkPart = doc.AddWorkbookPart ();
                wkPart.Workbook = new Workbook ();

                // Add WorksheetPart to the WorkbookPart
                WorksheetPart wsPart = wkPart.AddNewPart<WorksheetPart> ();
                SheetData sheetData = new SheetData ();

                object[][] matrix = new object[10][];
                for (int i = 0; i < 10; i++) {
                    matrix[i] = new object[] { "str_" + i, (i + 1) / 100d, DateTime.Now.AddMinutes (-i).ToString ("dd/MM/yyyy"), i * 100 };
                }

                int rowLength = 10;
                int cellLength = 5;
                for (uint i = 0; i < matrix.Length; i++) {
                    Row row = new Row { RowIndex = i + 1u };

                    for (int j = 0; j < matrix[i].Length; j++) {
                        CellValues data_type;
                        if (typeof (string) == matrix[i][j].GetType ()) data_type = CellValues.String;
                        else if (typeof (double) == matrix[i][j].GetType () || typeof (int) == matrix[i][j].GetType ()) data_type = CellValues.Number;
                        else if (typeof (DateTime) == matrix[i][j].GetType ()) data_type = CellValues.Date;
                        else data_type = CellValues.String;

                        Cell cell = new Cell {
                            CellReference = (char) (65 + j) + (1 + i).ToString (),
                            DataType = data_type, //CellValues.String,
                            CellValue = new CellValue (matrix[i][j].ToString ())
                        };
                        row.Append (cell);
                    }
                    sheetData.Append (row);

                }

                wsPart.Worksheet = new Worksheet (sheetData);

                // Add Sheets to the Workbook.
                Sheets sheets = doc.WorkbookPart.Workbook.AppendChild<Sheets> (new Sheets ());

                // Append a new sheet and associate it with the workbook.
                Sheet sheet = new Sheet ();
                sheet.Id = doc.WorkbookPart.GetIdOfPart (wsPart);
                sheet.SheetId = 1;
                sheet.Name = "sheet_1";
                sheets.Append (sheet);

                doc.Close ();

            }
        }

        public static void TesseractTest() {
            string testImagePath = "img.tif";
            try
            {
                 using (var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default))
                 {
                    using (var img = Pix.LoadFromFile(testImagePath))
                    {
                        using (var page = engine.Process(img))
                        {
                             var text = page.GetText();
                             System.Console.WriteLine(text);
                        }
                         
                    }    
                 }
            }
            catch (System.Exception)
            {
                
                throw;
            }


        }
    
    }
}