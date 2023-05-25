using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pruebasconexcell
{
    public class procesos
    {
        public static Dictionary<string, string> SeleccionarArchivo()
        {
            Dictionary<string, string> a = new Dictionary<string, string>();
            try
            {
                using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
                {
                    openFileDialog1.Multiselect = false;
                    openFileDialog1.Filter = "Sólo .xlsx (*.xlsx)|*.xlsx";
                    openFileDialog1.ShowDialog();
                    //var resultado = openFileDialog1.FileNames.Select(x => x).Union(openFileDialog1.SafeFileNames.Select(y=>y));
                    var resultado = openFileDialog1.FileNames.Zip(openFileDialog1.SafeFileNames, (x, y) => new { X = x, Y = y });
                    foreach (var ea in resultado)
                    {
                        string direccion = ea.X;
                        string nombre = ea.Y;
                        nombre = nombre.Substring(0, nombre.Length - 5);
                        a.Add(nombre, direccion);
                    }
                    
                }
                
                
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
            return a;
        }

        
        public static Dictionary<string, System.Data.DataTable> Abrir(string direccion)
        {

            Dictionary<string, System.Data.DataTable> tablas = new Dictionary<string, System.Data.DataTable>();
            
                using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(@direccion, false))
                {

                    WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                    IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();

                    
                    foreach (Sheet hoja in sheets)
                    {
                        string nombre = hoja.Name;
                        System.Data.DataTable tabla = new System.Data.DataTable();
                        string relationshipId = hoja.Id.Value;
                        WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                        Worksheet workSheet = worksheetPart.Worksheet;
                        SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                        IEnumerable<Row> rows = sheetData.Descendants<Row>();
                        foreach (Cell cell in rows.ElementAt(0))
                        {
                            tabla.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                        }

                    int ii = 0;//<- esto quizá es una chapuza
                    foreach (Cell cell in rows.ElementAt(1))
                    {
                        
                        tabla.Columns[ii].DataType = typeof(double);

                        var x = GetCellValue(spreadSheetDocument, cell);

                        //var z = int.TryParse(x.ToString(), out y) ? y : x;
                        //
                        if (double.TryParse(x.ToString(), out double y))
                        {
                            tabla.Columns[ii].DataType = typeof(double);
                        }
                        else
                        {
                            tabla.Columns[ii].DataType = typeof(string);
                        }
                        ii++;
                    }
                    
                    
                    foreach (Row row in rows)
                        {
                        DataRow tempRow;
                        //la primera fila son los nombres de las columnas, así que me la salto
                        if (row == rows.ElementAt(0))
                        {
                            continue;
                        } 
                        //comprueba que la línea no está vacía
                        if (row.Descendants<Cell>().Count() != 0)
                        {
                            tempRow = tabla.NewRow();


                            
                            int i = 0;
                            foreach (Cell cell in row)
                            {
                                if (cell.CellValue != null)
                                {
                                    var x = GetCellValue(spreadSheetDocument, cell);
                                    
                                    //var z = int.TryParse(x.ToString(), out y) ? y : x;
                                    if (double.TryParse(x.ToString(), out double y))
                                    {
                                        tempRow[i] = y;
                                    }
                                    else
                                    {
                                        if (tabla.Columns[i].DataType != typeof(string))
                                        {
                                            throw new Exception("Format souboru není správný, pro spravné fungovaní doporučuji použit neupravené soubory (.xlsx) z laboratorní statistiky OpenLIMS");
                                        }
                                        tempRow[i] = x;
                                    }
                                    
                                }

                                i++;
                            }

                            tabla.Rows.Add(tempRow);
                        }
                            
                           
                        }
                   
                    
                    
                    tablas.Add(nombre, tabla);
                    }

                }
            
            
            
            return tablas;
        }
        
        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            //if (cell.CellValue.InnerXml != null)
            //{
            if (cell.CellValue != null)
            {
                string value = cell.CellValue.InnerXml;
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
                }
                else
                {
                    return value;
                }
            }
            else
            {
                return " ";
            }
            //catch(Exception ex)
            //{
            //    //MessageBox.Show(ex.Message);

            //    return string.Empty;
            //    //error++;
            //}
                
            //}
            //else
            //{
            //    return string.Empty;
            //}
            
        }
    }
}
