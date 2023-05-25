using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Windows.Forms;
using Pruebasconexcell;
using raulexcell2.XLS;
using System.Data.Common;
using System.Collections.Generic;
using System.Collections;
using System.Linq.Expressions;
using tablitas;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Pruebasconexcell
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        static DataSet dsdatos = new DataSet("nombre");//Lo puse al principio, pero no vale para nada
        static DataTable dt = dsdatos.Tables.Add("tabla");
        static DataTable dt1 = dsdatos.Tables.Add("tablaauxiliar1");
        static DataTable dt2 = dsdatos.Tables.Add("tablaauxiliar2");
        static DataTable resultado = dsdatos.Tables.Add("tablaauxiliar3");
        //static DataTable tabla2 = dsdatos.Tables.Add("tabla2");
        Dictionary<string, string> archivos = new Dictionary<string, string>();

        int pasoint = 0;
        int countcol1 = 0;

        private void Form1_Load(object sender, EventArgs e)
        {
            label1.Text = string.Empty; // chapuza?

            

            //string imagePath = Path.Combine(Application.StartupPath, "ejemplo.png");
            //Image ejemplo = Image.FromFile(imagePath);
            //custommsg.showcustommsg("Návod k použiti", ejemplo);
        }

        private void Procesar_Click(object sender, EventArgs e)
        {
            if (dt.Columns.Count < 5) { return; }
            resultado = dt.Clone();
            resultado.Merge(dt);
            //Lleno con 0 los huecos vacíos
            foreach (DataRow dr in resultado.Rows)
            {
                foreach (DataColumn dc in resultado.Columns)
                {
                    if (string.IsNullOrEmpty(dr[dc.ColumnName].ToString()))
                    {
                        dr[dc.ColumnName] = 0;
                    }

                }
            }
             countcol1 = resultado.Columns.Count;

            

            for (int i = 4; i < countcol1; i++)
            {
                string n0 = resultado.Columns[i - 1].ColumnName;
                string n1 = resultado.Columns[i].ColumnName;
                resultado.Columns[i - 1].ColumnName = "uno";
                resultado.Columns[i].ColumnName = "dos";

                resultado.Columns.Add($"{n0} -> {n1}", typeof(double)/*, $"100*(( dos / uno ) - 1)"*/); ;//$"({resultado.Columns[i].ColumnName}/{resultado.Columns[i - 1].ColumnName}) - 1"
                foreach (DataRow dr in resultado.Rows)
                {
                    double rs = 100 * (((double)dr[i] / (double)dr[i - 1]) - 1);
                    dr[$"{n0} -> {n1}"] = rs;

                    if ((double)dr[i - 1]==0)
                    {
                        
                        
                        if ((double)dr[i] == 0)
                        {
                            dr[$"{n0} -> {n1}"] = 0;
                        }
                        else
                        {
                            dr[$"{n0} -> {n1}"] = 999;
                        }
                    }
                }

                resultado.Columns[i - 1].ColumnName = n0;
                resultado.Columns[i].ColumnName = n1;
            }

            resultado.Columns.Add("max/min", typeof(double));

            int countcol2 = resultado.Columns.Count;
            int countrows = resultado.Rows.Count;
            for (int i = 0; i< resultado.Rows.Count;)
            {
                ////para ver en qué momento falla:
                //string numero = resultado.Rows[i][0].ToString();
                //string metodo = resultado.Rows[i][1].ToString();
                ////--------------------------


                double max = resultado.Rows[i].ItemArray.Take(countcol1).Skip(3).Cast<double>().Max();
                double min = resultado.Rows[i].ItemArray.Take(countcol1).Skip(3).Cast<double>().Min();
                double porcentaje = double.TryParse(textBox1.Text, out double a) ? a : 0;
                if ((max / min) -1  < porcentaje/100 /*|| min == 0*/)
                {
                    resultado.Rows[i].Delete();
                    continue;
                    
                }
                
                

                double diferencia = Math.Round((max / min - 1)*100, 2);
                resultado.Rows[i][countcol2 - 1] = min == 0 ? 999 : diferencia;
                i++;
            }
            // convierto a int las columnas que siempre son números enteros para que salgan con el formato adecuado
            DataTable resultadocloned = resultado.Clone();
            for (int j = 0; j < countcol1; j++)
            {
                if (j == 1 || j==2) { continue; }
                resultadocloned.Columns[j].DataType = typeof(Int32);
            }
            foreach (DataRow row in resultado.Rows)
            {
                resultadocloned.ImportRow(row);
            }
            
            resultado = resultadocloned;


            dataGridView1.DataSource = resultado;
            dataGridView1.CellFormatting += (sender, e) =>
            {
                if (e.Value is double doubleValue)
                {
                    
                    e.Value = doubleValue.ToString("0.##");
                    e.FormattingApplied = true;

                    //if (e.ColumnIndex < countcol1-1)
                    //{
                    //    e.Value = doubleValue.ToString("0");
                    //    e.FormattingApplied = true;
                    //}
                }
            };



        }

        
        private void abrir_Click(object sender, EventArgs e)
        {
            // abre los archivos y junta las tablas directamente

            if (archivos.Count != 0)
            {
                foreach (var entrada in archivos)
                {
                    try
                    {
                        Dictionary<string, System.Data.DataTable> tablas = new Dictionary<string, System.Data.DataTable>();
                        tablas = procesos.Abrir(entrada.Value);
                        foreach (var DT in tablas)
                        {
                            string a = tablas.Count > 1 ? DT.Key : (textBox2.Text==string.Empty? entrada.Key : textBox2.Text);
                            if (dt.Rows.Count > 0)
                            {
                                if (!dt.Columns.Contains(a))
                                {
                                    dt1 = new DataTable();
                                    dt1.Merge(dt);

                                    dt2 = new DataTable();
                                    dt2 = DT.Value.DefaultView.ToTable(false, DT.Value.Columns[0].ColumnName, DT.Value.Columns[1].ColumnName, DT.Value.Columns[2].ColumnName, DT.Value.Columns[3].ColumnName) ;
                                    



                                    dt.Clear();
                                    dt.Merge(Unirtablas.añadircolumna(dt1, dt2, 0, 1, 3, a));

                                    //dt.Clear();
                                    //dt.Merge(dt3);
                                }
                                else
                                {
                                    MessageBox.Show($"Už existuje sloupec se jménem {entrada.Key}");
                                }
                            }
                            else
                            {
                                DT.Value.Columns[3].ColumnName = a;
                                dt2 = DT.Value.DefaultView.ToTable(false, DT.Value.Columns[0].ColumnName, DT.Value.Columns[1].ColumnName, DT.Value.Columns[2].ColumnName, DT.Value.Columns[3].ColumnName);
                                dt.Merge(dt2);
                                //string a = dt.Rows[0][0].GetType().Name;
                                //string str = string.Empty;
                                //for (int i = 0; i < dt.Columns.Count; i++)
                                //{
                                //    str += dt.Rows[0][i].GetType().Name;
                                //    str += "\n";
                                //}
                                //MessageBox.Show(str);
                            }
                            
                        }
                        

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                



            }
            else
            {
                MessageBox.Show("Není vybrán soubor");
            }
            dataGridView1.DataSource = dt;
            //DataTable joinedDataTable = result.CopyToDataTable();
            //*******************************

            // la columna 1 de dt es igual a la columna 0 de dt2, las junto usando eso.
            //la columna 2 de dt2 se copia a la última columna de dt

            //string a = (string)dt.Rows[0][0];
            //MessageBox.Show(a);
            //for (int i = 1;i < dt.Rows.Count; i++)
            //{
            //    if (dt.Rows[i][0] == dt.Rows[i - 1][0])
            //    {
            //        dt.Rows.RemoveAt(i);
            //    }
            //}
            //int b = dt.Rows.Count;
            //MessageBox.Show($"{a.ToString()} --> {b.ToString()}");



            /*DataColumn parentColumn = dsdatos.Tables["tabla"].Columns[0];
            DataColumn childColumn = dsdatos.Tables["tablaauxiliar"].Columns[0];
            DataRelation relation = new DataRelation("parent2Child", parentColumn, childColumn);
            dsdatos.Tables["tablaauxiliar"].ParentRelations.Add(relation);*/

            //dt.Columns.Add();
            //dt.Merge(dt2);
            //dt.AcceptChanges();

            //var results = from table1 in dt.AsEnumerable()
            //              join table2 in dt2.AsEnumerable() on (string)table1[0] equals (string)table2[0]
            //              select table1;
            //dt.Clear(); 
            //foreach (DataRow row in results)
            //{
            //    dt2.Rows.Add(row);
            //}





        }
        
        string archivo = string.Empty;
        private void guardar_Click(object sender, EventArgs e)
        {
            if (dt.Columns.Count == 0)
            {
                return;
            }


            //DataTable akkaka = (DataTable)dataGridView1.DataSource;


            archivo = manager.ExportaExcel((DataTable)dataGridView1.DataSource, "statistika");
            if (resultado.Columns.Count == 0) { return; }
            try
            {

                formato.formatear(archivo, resultado.Columns.Count, "00C6EFCE", ">10", 0u);
                formato.formatear(archivo, resultado.Columns.Count, "00FFC7CE", "<-10", 1u);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }


        }

        private void seleccionar_Click(object sender, EventArgs e)
        {
            Dictionary<string, string> aux = archivos;
            
            archivos = procesos.SeleccionarArchivo();
            if (archivos.Count == 0)
            {
                archivos = aux;
            }
            string a = string.Empty;
            string b = string.Empty;
            foreach (var s in archivos)
            {
                a += s.Value;
                //a += "\n";
                b += s.Key;
                //b += "\n";
            }
            label1.Text = b;
            textBox2.Text = b;
            //textBox1.Text = nombre;

                /* openFileDialog.InitialDirectory = "c:\\";
                 openFileDialog.Filter = "excel files (*.xslx)|*.xslx|";
                 openFileDialog.FilterIndex = 2;
                 openFileDialog.RestoreDirectory = true;

                 if (openFileDialog.ShowDialog() == DialogResult.OK)
                 {
                     //Get the path of specified file
                     archivo  = openFileDialog.FileName;

                     //Read the contents of the file into a stream

                 }*/
            //}
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
            if ((e.KeyChar == (char)13))
            {
                Procesar_Click(sender, e);
            }
        }

        private void paso_Click(object sender, EventArgs e)
        {
            if (pasoint == 0)
            {
                textBox3.Text = "2.   Vybrat nejmenší rozdíl co považujeme za významný (automaticky přednastaveno 10 %) a tlačit Zpracovat.\r\n   • Pokud je potřeba zobrazit kompletní statistiku, musí se zvolit rozdíl 0 %.";
            }
            if (pasoint == 1)
            {
                textBox3.Text = "3.   Zpracovaná data lze uložit jako excelový soubor tlačítkem Uložit.\r\n   • Lze změnit procento, ale je nutné uložit jednu tabulku za druhou.";
            }


            pasoint ++;
        }
    }
}