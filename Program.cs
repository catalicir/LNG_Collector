using System;
using System.Collections.Generic;
using System.Text.Json;
using Newtonsoft.Json;
using System.Net;
using System.IO;
using System.Threading.Tasks;
using _LNG_Collector.Utils;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Linq;
using NPOI.HSSF.UserModel;

namespace _LNG_Collector
{
    internal class LNG_Situatii_Zilnice
    {
        public string settings_file = "settings.json";
        
        static void Main(string[] args)
        {
            //Pentru real-time environment
            //string data_curenta = DateTime.Now.ToString("dd-MM-yyyy");
            //string folder_curent = Directory.GetCurrentDirectory();


            //Pentru testare
            string data_curenta = "01-12-2022";
            string folder_curent = @"C:\GitHub\LNG_Collector\";

            Console.WriteLine("Initiere Generare Situatii Zilnice LNG: " + data_curenta);
            Console.WriteLine("...");
            
            //citeste fisierele de intrare
            Console.WriteLine("Verifica existenta date de intrare...");
            CitesteSiVerificaFisierele(folder_curent);

            //Creaza structura pentru data curenta
            Directory.CreateDirectory(folder_curent+@"\"+data_curenta);
            Directory.CreateDirectory(folder_curent + @"\" + data_curenta +@"\input");
            Directory.CreateDirectory(folder_curent + @"\" + data_curenta +@"\output");

            //copiaza fisierele de intrare in directorul zilnic de input
            
            //copiaza fisierele template pentru a fi umplute cu dare in diectorul zilnic de output
            CopiazaFisiere(folder_curent + @"\templates\", folder_curent + @"\" + data_curenta + @"\output\");
            CopiazaFisiere(folder_curent + @"\transfer_input\", folder_curent + @"\" + data_curenta + @"\input\");
            //redenumeste fisiere din output - trebuie?


            //copiaza continut din input in output
            //CopiazaContinut(@"\" + data_curenta + @"\input\", @"\" + data_curenta + @"\output\", "template 1", "A2", "C2");
            CopiazaContinutTest(folder_curent + @"\" + data_curenta + @"\input\input 11.xlsx", folder_curent + @"\" + data_curenta + @"\output\template 01.xlsx");
            //order information on output sheets


            //refresh pivot tables
            
            

        }

        private static void CopiazaContinutTest(string inputFilePath, string outputFilePath)
        {
            DataTable dtTable = new DataTable();
            var fs = new FileStream(inputFilePath, FileMode.Open);
            try
            {
                //copiaza din input 11.xls in template 01.xls:
                //  sheet "Input 1", cols A, B de la row 2 incolo in "template 1", cols A, B de la row 2
                //  sheet "Input 1", cols C de la row 2 incolo in ""template 1", cols C de la row 2 incolo

                //creaza file streams penreu cele 2 fisiere
                
                List<string> rowList = new List<string>();
                
                using (fs)
                {
                    dtTable = GetDataFromFileWorksheet(dtTable, rowList, fs, 0);
                }
                //pentru verificare
                string whatIveGotSoFar = JsonConvert.SerializeObject(dtTable);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Copierea continutului a esuat: ");
                Console.WriteLine(ex.ToString());
            }

            try
            {
                //deschide fisierul destinatie pentru scriere
                WriteDataToFileWorksheet(dtTable, outputFilePath, 0);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Scrierea continutului a esuat: ");
                Console.WriteLine(ex.ToString());
            }
            finally { fs.Close(); }
        }

        private static void WriteDataToFileWorksheet(DataTable table, string outputFilePath, int worksheetno)
        {
            //open destination file to write
            using (var fs = new FileStream(outputFilePath, FileMode.Open, FileAccess.ReadWrite))
            {
                fs.Position = 0;
                XSSFWorkbook workbook = new XSSFWorkbook(fs);
                //workbook.Write(fs, true);
                var excelSheet = workbook.GetSheetAt(worksheetno);

                //List<DataColumn> columns = new List<DataColumn>();
                IRow row = excelSheet.GetRow(1);
                int columnIndex = 0;

                // you can create also column header - comment for now
                /*foreach (System.Data.DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);
                    row.CreateCell(columnIndex).SetCellValue(column.ColumnName);
                    columnIndex++;
                }*/

                int rowIndex = 1;
                foreach (DataRow dsrow in table.Rows)
                {
                    //row = excelSheet.CreateRow(rowIndex);
                    int cellIndex = 0;
                    foreach (DataColumn col in table.Columns)
                    {
                        if (row == null) { 
                            row = excelSheet.CreateRow(rowIndex++);
                        }
                        row.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
                        cellIndex++;
                    }

                    rowIndex++;
                }

                workbook.Write(fs, false);
                fs.Close(); 
            }
        }

        private static DataTable GetDataFromFileWorksheet(DataTable dtTable, List<string> rowList, FileStream fs, int worksheetno)
        {
            IWorkbook workbook = new XSSFWorkbook(fs);
            var sheet = workbook.GetSheetAt(worksheetno);
            IRow headerRow = sheet.GetRow(0);
            int cellCount = headerRow.LastCellNum;
            for (int j = 0; j < cellCount; j++)
            {
                ICell cell = headerRow.GetCell(j);
                if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
                {
                    dtTable.Columns.Add(cell.ToString());
                }
            }
            for (int i = (sheet.FirstRowNum); i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue;
                if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;
                for (int j = row.FirstCellNum; j < cellCount; j++)
                {
                    if (row.GetCell(j) != null)
                    {
                        if (!string.IsNullOrEmpty(row.GetCell(j).ToString()) && !string.IsNullOrWhiteSpace(row.GetCell(j).ToString()))
                        {
                            rowList.Add(row.GetCell(j).ToString());
                        }
                    }
                }
                if (rowList.Count > 0)
                    dtTable.Rows.Add(rowList.ToArray());
                rowList.Clear();
            }
            fs.Close();

            return dtTable;
        }


        private static void CopiazaFisiere(string inputPath, string dailyInput)
        {
            foreach (var newPath in Directory.GetFiles(inputPath, "*.*", SearchOption.AllDirectories))
            {
                File.Copy(newPath, newPath.Replace(inputPath, dailyInput));
                Console.WriteLine(newPath+": Fisier copiat in " + dailyInput);
            }
        }

        private static void CitesteSiVerificaFisierele(string folder_curent)
        {
            //citeste din setari ce fisiere trebuie sa existe
            Setari setari = new Setari();
            try
            {
                setari = JsonConvert.DeserializeObject<Setari>(File.ReadAllText(folder_curent + @"\settings.json"));
            }catch (Exception ex)
            {
                Console.WriteLine("Fisier de setari eronat" + ex);
                return;
            }
            
            //
            foreach (string filename in setari.InputFiles)
            {
                //verify all files in the list
                bool allOK = true;
                if (File.Exists(folder_curent + @"\transfer_input\" + filename))
                {
                    Console.WriteLine(folder_curent + @"\transfer_input\" + filename + " -> exista. OK");
                }
                else { 
                    allOK = false;
                    Console.WriteLine(folder_curent + @"\transfer_input\" + filename + " -> NU EXISTA!. NOK");
                }
                if (!allOK) {
                    Console.WriteLine("Nu toate fisierele de input sunt prezente");
                    Console.WriteLine("Datele zilnice nu au fost inca incarcate. Incercati mai tarziu sau verificati folderul input din data curenta");
                    return;
                }
            }
        }

        
    }
}
