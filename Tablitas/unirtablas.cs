using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pruebasconexcell
{
    public static class Unirtablas
    {
        /// <summary>
        /// añade una columna al final de una tabla1 desde una tabla2 teniendo ambas tablas una columna en común
        /// </summary>
        /// <param name="tabla1">tabla que recibe</param>
        /// <param name="tabla2">tabla que aporta</param>
        /// <param name="clave1">Columna en común 1</param>
        /// <param name="clave2">Columna en común 2</param>
        /// <param name="columna">indice de la columna a aportar</param>
        /// <param name="nombre">nombre de la columna a aportar</param>
        /// <returns></returns>
        public static DataTable añadircolumna(DataTable tabla1, DataTable tabla2, int clave1, int clave2, int columna, string nombre)
        {
            int count = tabla1.Columns.Count;
            //Añado una columna a la tabla que recibe y le doy el mismo nombre a la columna que recibe y a la que aporta
            tabla2.Columns[columna].ColumnName = nombre;
            tabla1.Columns.Add(nombre, typeof(double));
            //foreach (DataRow dr in tabla1.Rows)
            //{
            //    dr[count] = 0;
            //}
            //Ordeno las columnas para agilizar el proceso
            tabla2.DefaultView.Sort = tabla2.Columns[clave1].ColumnName;
            tabla2 = tabla2.DefaultView.ToTable();
            tabla1.DefaultView.Sort = tabla1.Columns[clave1].ColumnName;
            tabla1 = tabla1.DefaultView.ToTable();

            for (int i = 0; i < tabla1.Rows.Count; i++)
            {
                for (int j = 0; j < tabla2.Rows.Count; j++)
                {
                    string a = tabla1.Rows[i][clave1].ToString();
                    string b = tabla2.Rows[j][clave1].ToString();
                    string c = tabla1.Rows[i][clave2].ToString();
                    string d = tabla2.Rows[j][clave2].ToString();
                    if (a == b & c == d)
                    {
                        tabla1.Rows[i][count] = tabla2.Rows[j][columna];
                        tabla2.Rows.RemoveAt(j);
                        break;
                    }
                }
            }
            tabla1.Merge(tabla2);
            tabla1.DefaultView.Sort = tabla1.Columns[0].ColumnName;
            tabla1 = tabla1.DefaultView.ToTable();
            //foreach (DataRow dr in tabla1.Rows)
            //{
            //    if (string.IsNullOrEmpty(dr[count - 1].ToString()))
            //    {
            //        dr[count - 1] = 0;
            //    }
            //}

            return tabla1;
        }
    }
}
