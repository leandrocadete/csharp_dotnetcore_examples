using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace csharp_dotnetcore_examples {
    class Program {
        static void Main (string[] args) {

            //ReadExcel("teste1.xlsx");
            string filePath = "por.traineddata";
            using (System.IO.FileStream fs = new System.IO.FileStream(filePath, System.IO.FileMode.Open))
            {
                if(fs.Length % 4 != 0) { 
                    Int64 qtds = fs.Length;
                    byte[] buffer = new byte[qtds];
                    fs.Read(buffer, 0, (int) qtds);
                    string strOut = System.Convert.ToBase64String(buffer);

                    int length = strOut.Length/4;
                    int mod4 = strOut.Length % 4;
                    System.Console.WriteLine("length => {0} = {1}/{2}", length, strOut.Length, 4);
                    for (int i = 0; i < 4; i++)
                    {
                        System.Console.WriteLine("substring {0}, {1}", i*length, (i+1) * length);
                        if(i == 3) {
                            System.IO.File.WriteAllText($"output_0{i+1}.txt", strOut.Substring(i * length));    
                            continue;
                        }
                        System.IO.File.WriteAllText($"output_0{i+1}.txt", strOut.Substring(i * length, length));
                    }
                } else {
                    System.Console.WriteLine(fs.Length);
                }
            }
        }

        public static void ReadExcel (string path) {

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open (path, true)) {
                Sheets sheets = doc.WorkbookPart.Workbook.Sheets;
                WorkbookPart wPart = doc.WorkbookPart;

                IEnumerable<Sheet> shs = sheets.ChildElements.Cast<Sheet> ();
                var sheet1 = shs.FirstOrDefault<Sheet> (s => s.Name == "Planilha1"); // get tab by name

                Worksheet workSheet = ((WorksheetPart) wPart.GetPartById (sheet1.Id)).Worksheet;
                List<SheetData> rows = workSheet.ChildElements.OfType<SheetData> ().ToList ();

                string currCellValue = null;

                List<List<string>> lstSheet = new List<List<string>> (rows[0].ChildElements.Count);

                for (int i = 0; i < rows[0].ChildElements.Count; i++) {
                    lstSheet.Add (new List<string> (4));

                    Row currentrow = (Row) rows[0].ChildElements.GetItem (i);

                    Cell[] cells = new Cell[] {
                        (Cell) currentrow.ChildElements.GetItem (0)
                        ,(Cell) currentrow.ChildElements.GetItem (1)
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
    }
}