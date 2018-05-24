using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NormalizarCP.Entidades;

namespace NormalizarCP.Datos
{
    public class CalleDAO
    {
        #region "PATRON SINGLETON"
        private static CalleDAO dao = null;
        private CalleDAO() { }
        public static CalleDAO getInstance()
        {
            if (dao == null)
            {
                dao = new CalleDAO();
            }
            return dao;
        }
        #endregion 

        /// <summary>
        /// Lee todos los registros del archivo de calles + zonas
        /// </summary>
        /// <param name="lista"></param>
        /// <returns></returns>
        public static List<Calle> readCalles(List<Calle> lista)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\Mapas\calles-zonas.xlsx");
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;


            for (int fila = 2; fila <= rowCount; fila++)
            {
                Console.WriteLine(fila);
                Calle calle = new Calle();
                try
                {
                    calle.id = Convert.ToInt32(xlRange.Cells[fila, 12].Value2.ToString());
                }
                catch (Exception)
                {
                    continue;
                }
                try
                {
                    calle.nro_zona = xlRange.Cells[fila, 3].Value2.ToString();
                }
                catch (Exception)
                {
                    continue;
                }
                
                try
                {
                    calle.calle = xlRange.Cells[fila, 14].Value2.ToString();
                }
                catch (Exception)
                {
                    calle.calle = "";
                }
                try
                {
                    calle.altura_ini = Convert.ToInt32(xlRange.Cells[fila, 15].Value2.ToString());
                }
                catch (Exception)
                {
                    calle.altura_ini = 0;
                }
                
                lista.Add(calle);
            }
            xlApp.DisplayAlerts = false;
            xlWorkbook.Close();
            xlApp.Quit();
            return lista;
        }


        /// <summary>
        /// Copia todos los registros scrapeados en el archivo indicado
        /// </summary>
        /// <param name="list"></param>
        /// <param name="nombreArchivo"></param>
        public static void cpsToExcel(List<Calle> list, string nombreArchivo)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }

            xlApp.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook wb = xlApp.Workbooks.Open(@"D:\Mapas\" + nombreArchivo);
            Microsoft.Office.Interop.Excel._Worksheet ws = wb.Sheets[1];

            if (ws == null)
            {
                Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
            }


            Microsoft.Office.Interop.Excel.Range xlRange = ws.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            int fila = rowCount + 1;

            foreach (Calle calle in list)
            {
                ws.Cells[fila, 1] = calle.id;
                ws.Cells[fila, 2] = calle.nro_zona;
                ws.Cells[fila, 3] = calle.calle;
                ws.Cells[fila, 4] = calle.altura_ini;
                ws.Cells[fila, 5] = calle.cp;
                fila++;
            }

            ws.Rows.WrapText = false;
            ws.Columns.WrapText = false;
            xlApp.DisplayAlerts = false;
            wb.SaveAs(@"D:\Mapas\" + nombreArchivo, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            wb.Close();
            xlApp.Quit();
        }

        /// <summary>
        /// Copia un registro scrapeado en el archivo indicado
        /// </summary>
        /// <param name="list"></param>
        /// <param name="nombreArchivo"></param>
        public static void cpToExcel(Calle calle, string nombreArchivo)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();


            if (xlApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }

            xlApp.Visible = false;
            
            Microsoft.Office.Interop.Excel.Workbook wb = xlApp.Workbooks.Open(@"D:\Mapas\" + nombreArchivo);
            Microsoft.Office.Interop.Excel._Worksheet ws = wb.Sheets[1];

            if (ws == null)
            {
                Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
            }

            Microsoft.Office.Interop.Excel.Range xlRange = ws.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            int fila = rowCount + 1;

            ws.Cells[fila, 1] = calle.id;
            ws.Cells[fila, 2] = calle.nro_zona;
            ws.Cells[fila, 3] = calle.calle;
            ws.Cells[fila, 4] = calle.altura_ini;
            ws.Cells[fila, 5] = calle.cp;
            fila++;

            ws.Rows.WrapText = false;
            ws.Columns.WrapText = false;
            xlApp.DisplayAlerts = false;
            wb.SaveAs(@"D:\Mapas\" + nombreArchivo, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            wb.Close();
            xlApp.Quit();
        }

    }
}
