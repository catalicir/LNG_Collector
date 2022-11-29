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
using NPOI.SS.Formula.Functions;
using System.Drawing;
using NPOI.XWPF.UserModel;
using ICell = NPOI.SS.UserModel.ICell;
using static System.Net.WebRequestMethods;
using File = System.IO.File;
using NPOI.XSSF.UserModel.Charts;

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
            Directory.CreateDirectory(folder_curent + @"\" + data_curenta);
            Directory.CreateDirectory(folder_curent + @"\" + data_curenta + @"\input");
            Directory.CreateDirectory(folder_curent + @"\" + data_curenta + @"\output");

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

                List<object> rowList = new List<object>();

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
                WriteDataToFileWorksheet(dtTable, outputFilePath, 0, 0);
                WriteDataToFileWorksheet(dtTable, outputFilePath, 1, 1);
                //WriteExcel(dtTable);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Scrierea continutului a esuat: ");
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                fs.Close();
            }
        }

        static void WriteExcel(DataTable dtTable)
        {
            List<UserDetails> persons = new List<UserDetails>()
            {
                new UserDetails() {ID="1001", Name="ABCD", City ="City1", Country="USA"},
                new UserDetails() {ID="1002", Name="PQRS", City ="City2", Country="INDIA"},
                new UserDetails() {ID="1003", Name="XYZZ", City ="City3", Country="CHINA"},
                new UserDetails() {ID="1004", Name="LMNO", City ="City4", Country="UK"},
           };

            // Lets converts our object data to Datatable for a simplified logic.
            // Datatable is most easy way to deal with complex datatypes for easy reading and formatting.

            DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(persons), (typeof(DataTable)));
            var memoryStream = new MemoryStream();

            using (var fs = new FileStream(@"C:\GitHub\LNG_Collector\01-12-2022\output\template 01.xlsx", FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook();
                //ISheet excelSheet = workbook.CreateSheet("Sheet1");
                ISheet excelSheet = workbook.GetSheetAt(0);


                List<String> columns = new List<string>();
                IRow row = excelSheet.CreateRow(0);
                int columnIndex = 0;

                foreach (System.Data.DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);
                    row.CreateCell(columnIndex).SetCellValue(column.ColumnName);
                    columnIndex++;
                }

                int rowIndex = 1;
                foreach (DataRow dsrow in table.Rows)
                {
                    row = excelSheet.CreateRow(rowIndex);
                    int cellIndex = 0;
                    foreach (String col in columns)
                    {
                        row.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
                        cellIndex++;
                    }

                    rowIndex++;
                }
                fs.Close();
                FileStream file = new FileStream(@"C:\GitHub\LNG_Collector\01-12-2022\output\template 01.xlsx", FileMode.Create);
                workbook.Write(file, false);
            }

        }

        private static void WriteDataToFileWorksheet(DataTable table, string outputFilePath, int worksheetno, int colIndextoWrite)
        {
            //open destination file to write
            var fs = new FileStream(outputFilePath, FileMode.Open, FileAccess.Read);

            fs.Position = 0;
            XSSFWorkbook workbook = new XSSFWorkbook(fs);
            fs.Close();
            //workbook.Write(fs, false);
            var excelSheet = workbook.GetSheetAt(worksheetno);

            //Formateaza pentru diverse tipuri de date:
            ICellStyle _doubleCellStyle = workbook.CreateCellStyle();
            _doubleCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0.###");

            ICellStyle _intCellStyle = workbook.CreateCellStyle();
            _intCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0");

            ICellStyle _boolCellStyle = workbook.CreateCellStyle();
            _boolCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("BOOLEAN");

            ICellStyle _dateCellStyle = workbook.CreateCellStyle();
            _dateCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("dd-MM-yyyy");

            ICellStyle _dateTimeCellStyle = workbook.CreateCellStyle();
            _dateTimeCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("dd-MM-yyyy HH:mm:ss");

            int rowIndex = 0;
            foreach (DataRow dsrow in table.Rows)
            {
                //skip header
                if (rowIndex != 0)
                {
                    
                    int cellIndex = colIndextoWrite;
                    //ia linia si daca nu exista o creaza - daca exista, inseamna ca mai am info pe linia respectiva pe care nu vreau sa le suprascriu 
                    IRow row = excelSheet.GetRow(rowIndex);
                    foreach (DataColumn col in table.Columns)
                    {
                        if (row == null)
                        {
                            row = excelSheet.CreateRow(rowIndex);
                        }

                        ICell cell = null; //<- cell curent                      
                        object cellValue = dsrow[col]; //<- valoarea curenta a cell
                        
                        /*switch (cellValue.GetType().FullName)
                        {
                            case "System.Boolean":
                                if (cellValue != DBNull.Value)
                                {
                                    cell = row.CreateCell(cellIndex, CellType.Boolean);

                                    if (Convert.ToBoolean(cellValue)) { cell.SetCellFormula("TRUE()"); }
                                    else { cell.SetCellFormula("FALSE()"); }

                                    cell.CellStyle = _boolCellStyle;
                                }
                                break;

                            case "System.String":
                                if (cellValue != DBNull.Value)
                                {
                                    cell = row.CreateCell(cellIndex, CellType.String);
                                    cell.SetCellValue(Convert.ToString(cellValue));
                                }
                                break;

                            case "System.Int32":
                                if (cellValue != DBNull.Value)
                                {
                                    cell = row.CreateCell(cellIndex, CellType.Numeric);
                                    cell.SetCellValue(Convert.ToInt32(cellValue));
                                    cell.CellStyle = _intCellStyle;
                                }
                                break;
                            case "System.Int64":
                                if (cellValue != DBNull.Value)
                                {
                                    cell = row.CreateCell(cellIndex, CellType.Numeric);
                                    cell.SetCellValue(Convert.ToInt64(cellValue));
                                    cell.CellStyle = _intCellStyle;
                                }
                                break;
                            case "System.Decimal":
                                if (cellValue != DBNull.Value)
                                {
                                    cell = row.CreateCell(cellIndex, CellType.Numeric);
                                    cell.SetCellValue(Convert.ToDouble(cellValue));
                                    cell.CellStyle = _doubleCellStyle;
                                }
                                break;
                            case "System.Double":
                                if (cellValue != DBNull.Value)
                                {
                                    cell = row.CreateCell(cellIndex, CellType.Numeric);
                                    cell.SetCellValue(Convert.ToDouble(cellValue));
                                    cell.CellStyle = _doubleCellStyle;
                                }
                                break;

                            case "System.DateTime":
                                if (cellValue != DBNull.Value)
                                {
                                    cell = row.CreateCell(cellIndex, CellType.Numeric);
                                    cell.SetCellValue(Convert.ToDateTime(cellValue));

                                    //Si No tiene valor de Hora, usar formato dd-MM-yyyy
                                    DateTime cDate = Convert.ToDateTime(cellValue);
                                    if (cDate != null && cDate.Hour > 0) { cell.CellStyle = _dateTimeCellStyle; }
                                    else { cell.CellStyle = _dateCellStyle; }
                                }
                                break;
                            default:
                                break;
                        }*/

                        row.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
                        cellIndex++;
                    }
                  
                }
                rowIndex++;
            }

            // Trebuie pus create chiar daca folosesc acelasi fisier
            fs = new FileStream(outputFilePath, FileMode.Create, FileAccess.Write);
            workbook.Write(fs, false);
            workbook.SetForceFormulaRecalculation(true);
            XSSFFormulaEvaluator.EvaluateAllFormulaCells(workbook);
            
            workbook.Close();
            fs.Close();

        }

        private static DataTable GetDataFromFileWorksheet(DataTable dtTable, List<object> rowList, FileStream fs, int worksheetno)
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
                            var cellValue = row.GetCell(j);
                            

                            /*switch (cellValue.CellType){
                                case CellType.String:
                                    rowList.Add(cellValue.ToString());
                                    break;
                                case CellType.Numeric:
                                    rowList.Add(cellValue.NumericCellValue);
                                    break;
                                case CellType.Boolean:
                                    rowList.Add(cellValue.BooleanCellValue);
                                    break;
                                case CellType.Formula:
                                    rowList.Add(cellValue.CellFormula);
                                    break;
                                default:
                                    rowList.Add(cellValue.ToString());
                                    break;
                            }  */

                            //nu stiu cum sa fac cu cele de tip Data.
                            rowList.Add(cellValue.ToString());


                        }
                    }
                }
                if (rowList.Count > 0)
                    dtTable.Rows.Add(rowList.ToArray());
                rowList.Clear();
            }
            workbook.Close();
            fs.Close();

            return dtTable;
        }


        private static void CopiazaFisiere(string inputPath, string dailyInput)
        {
            foreach (var newPath in Directory.GetFiles(inputPath, "*.*", SearchOption.AllDirectories))
            {
                File.Copy(newPath, newPath.Replace(inputPath, dailyInput));
                Console.WriteLine(newPath + ": Fisier copiat in " + dailyInput);
            }
        }

        private static void CitesteSiVerificaFisierele(string folder_curent)
        {
            //citeste din setari ce fisiere trebuie sa existe
            Setari setari = new Setari();
            try
            {
                setari = JsonConvert.DeserializeObject<Setari>(File.ReadAllText(folder_curent + @"\settings.json"));
            }
            catch (Exception ex)
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
                else
                {
                    allOK = false;
                    Console.WriteLine(folder_curent + @"\transfer_input\" + filename + " -> NU EXISTA!. NOK");
                }
                if (!allOK)
                {
                    Console.WriteLine("Nu toate fisierele de input sunt prezente");
                    Console.WriteLine("Datele zilnice nu au fost inca incarcate. Incercati mai tarziu sau verificati folderul input din data curenta");
                    return;
                }
            }
        }


        private void DataTable_To_Excel(DataTable pDatos, string pFilePath)
        {
            try
            {
                if (pDatos != null && pDatos.Rows.Count > 0)
                {
                    IWorkbook workbook = null;
                    ISheet worksheet = null;

                    using (FileStream stream = new FileStream(pFilePath, FileMode.Create, FileAccess.ReadWrite))
                    {
                        string Ext = System.IO.Path.GetExtension(pFilePath); //<-Extension del archivo
                        switch (Ext.ToLower())
                        {
                            case ".xls":
                                HSSFWorkbook workbookH = new HSSFWorkbook();
                                NPOI.HPSF.DocumentSummaryInformation dsi = NPOI.HPSF.PropertySetFactory.CreateDocumentSummaryInformation();
                                dsi.Company = "Cutcsa"; dsi.Manager = "Departamento Informatico";
                                workbookH.DocumentSummaryInformation = dsi;
                                workbook = workbookH;
                                break;

                            case ".xlsx": workbook = new XSSFWorkbook(); break;
                        }

                        worksheet = workbook.CreateSheet(pDatos.TableName); //<-Usa el nombre de la tabla como nombre de la Hoja

                        //CREAR EN LA PRIMERA FILA LOS TITULOS DE LAS COLUMNAS
                        int iRow = 0;
                        if (pDatos.Columns.Count > 0)
                        {
                            int iCol = 0;
                            IRow fila = worksheet.CreateRow(iRow);
                            foreach (DataColumn columna in pDatos.Columns)
                            {
                                ICell cell = fila.CreateCell(iCol, CellType.String);
                                cell.SetCellValue(columna.ColumnName);
                                iCol++;
                            }
                            iRow++;
                        }

                        //FORMATOS PARA CIERTOS TIPOS DE DATOS
                        ICellStyle _doubleCellStyle = workbook.CreateCellStyle();
                        _doubleCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0.###");

                        ICellStyle _intCellStyle = workbook.CreateCellStyle();
                        _intCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0");

                        ICellStyle _boolCellStyle = workbook.CreateCellStyle();
                        _boolCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("BOOLEAN");

                        ICellStyle _dateCellStyle = workbook.CreateCellStyle();
                        _dateCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("dd-MM-yyyy");

                        ICellStyle _dateTimeCellStyle = workbook.CreateCellStyle();
                        _dateTimeCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("dd-MM-yyyy HH:mm:ss");

                        //AHORA CREAR UNA FILA POR CADA REGISTRO DE LA TABLA
                        foreach (DataRow row in pDatos.Rows)
                        {
                            IRow fila = worksheet.CreateRow(iRow);
                            int iCol = 0;
                            foreach (DataColumn column in pDatos.Columns)
                            {
                                ICell cell = null; //<-Representa la celda actual                               
                                object cellValue = row[iCol]; //<- El valor actual de la celda

                                switch (column.DataType.ToString())
                                {
                                    case "System.Boolean":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Boolean);

                                            if (Convert.ToBoolean(cellValue)) { cell.SetCellFormula("TRUE()"); }
                                            else { cell.SetCellFormula("FALSE()"); }

                                            cell.CellStyle = _boolCellStyle;
                                        }
                                        break;

                                    case "System.String":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.String);
                                            cell.SetCellValue(Convert.ToString(cellValue));
                                        }
                                        break;

                                    case "System.Int32":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToInt32(cellValue));
                                            cell.CellStyle = _intCellStyle;
                                        }
                                        break;
                                    case "System.Int64":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToInt64(cellValue));
                                            cell.CellStyle = _intCellStyle;
                                        }
                                        break;
                                    case "System.Decimal":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToDouble(cellValue));
                                            cell.CellStyle = _doubleCellStyle;
                                        }
                                        break;
                                    case "System.Double":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToDouble(cellValue));
                                            cell.CellStyle = _doubleCellStyle;
                                        }
                                        break;

                                    case "System.DateTime":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToDateTime(cellValue));

                                            //Si No tiene valor de Hora, usar formato dd-MM-yyyy
                                            DateTime cDate = Convert.ToDateTime(cellValue);
                                            if (cDate != null && cDate.Hour > 0) { cell.CellStyle = _dateTimeCellStyle; }
                                            else { cell.CellStyle = _dateCellStyle; }
                                        }
                                        break;
                                    default:
                                        break;
                                }
                                iCol++;
                            }
                            iRow++;
                        }

                        workbook.Write(stream, false);
                        stream.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }



}
