using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Exporta
{
    public class Exporta
    {
        /// <summary>
        /// Método que exporta a un fichero Excel el contenido de un DataGridView
        /// </summary>
        /// <param name="grd">DataGridView que contiene los datos a exportar</param>    
        /// <param name="fichero">SaveFileDialog donde se guardara el archivo</param>    
        public static bool ExportarDataGridViewExcelXML(DataGridView grd, SaveFileDialog fichero)
        {
            try
            {
                // Obtenermos una lsita generica con la informacion del DataGridView
                var Result = grd.Rows.OfType<DataGridViewRow>().Select(
                r => r.Cells.OfType<DataGridViewCell>().Select(c => c.Value).ToArray()).ToList();

                // Creamos la ruta para el archivo temporal con el que trabajaremos
                var archTemp = string.Format("{0}\\{1}.tmp", Path.GetDirectoryName(fichero.FileName), Path.GetFileNameWithoutExtension(fichero.FileName));

                // Abrimos el archivo temporal
                using (StreamWriter mylogs = new StreamWriter(archTemp, false, Encoding.GetEncoding(1252)))
                {
                    string fila = "";

                    // Recorremos el DataGridView y llenamos el encabezado en el archivo temporal
                    for (int i = 1; i <= grd.Columns.Count; i++)
                    {
                        if (i == 1)
                            fila += string.Format("{0}", grd.Columns[i - 1].HeaderText);
                        else
                            fila += string.Format(",{0}", grd.Columns[i - 1].HeaderText);
                    }
                    mylogs.WriteLine(fila);

                    // Recorremos fila por fila para obtener sus columnas
                    Result.ForEach(line =>
                    {
                        // Limpiamos
                        fila = string.Empty;
                        // Creamos lista con los campos de la fila
                        var campos = line.Select(x => x.ToString()).ToList();
                        // Recorremos los campos de la fila
                        campos.ForEach(campo =>
                        {
                            // Validamos si el tipo de datos del campo es DateTime
                            DateTime fecha;
                            if (DateTime.TryParse(campo.ToString(), out fecha))
                            {
                                // Actualizamos el campo con la fecha que tiene la cultura correcta(Esto 
                                // es devido a que la libreria Microsoft.Office.Interop.Excel con la que
                                // exportaremos la informacion a excel, solo funciona con la cultura "en-US")
                                campo = fecha.ToString("d", CultureInfo.CreateSpecificCulture("en-US"));
                            }

                            // Remplazamos las compas ya que este sera el separador para cada columna en 
                            // el archivo temporal y las fechas que no quiero mostrar (en mi caso muy particular)
                            campo = campo.ToString().Replace(",", " ").Replace("1/1/0001", "");
                            // Vamos creando la cadena para llenar el archivo.
                            fila += string.Format("{0},", campo.ToString());
                        });
                        // Escribimos en el archivo la cadena resultado del recorrido de la fila
                        mylogs.WriteLine(fila);
                    });
                    // Cerramos el archivo (No deveria de ser necesario ya que lo estamos manejando en USING, 
                    // pero me ha presentado errores, por lo que opte por meterlo.)
                    mylogs.Close();
                }

                // Instanciamos la libreria 
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                // Abrimos el archivo temporal
                Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Open(archTemp, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                // Guardamos el archivo en formato xls
                wb.SaveAs(fichero.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlXMLSpreadsheet, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                // Cerramos el archivo temporal
                wb.Close();
                app.Quit();

                // Eliminamos el archivo temporal
                File.Delete(archTemp);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }        
    }
}
