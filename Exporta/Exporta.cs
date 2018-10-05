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
        public static bool ExportarDataGridViewExcelXML(DataGridView grd, SaveFileDialog fichero)
        {
            try
            {
                var Result = grd.Rows.OfType<DataGridViewRow>().Select(
                r => r.Cells.OfType<DataGridViewCell>().Select(c => c.Value).ToArray()).ToList();

                var archTemp = string.Format("{0}\\{1}.tmp", Path.GetDirectoryName(fichero.FileName), Path.GetFileNameWithoutExtension(fichero.FileName));

                using (StreamWriter mylogs = new StreamWriter(archTemp, false, Encoding.GetEncoding(1252)))
                {
                    string fila = "";

                    for (int i = 1; i <= grd.Columns.Count; i++)
                    {
                        if (i == 1)
                            fila += string.Format("{0}", grd.Columns[i - 1].HeaderText);
                        else
                            fila += string.Format(",{0}", grd.Columns[i - 1].HeaderText);
                    }

                    mylogs.WriteLine(fila);

                    Result.ForEach(line =>
                    {
                        fila = string.Join(",", line.Select(x => x.ToString()).ToArray());
                        mylogs.WriteLine(fila);
                    });
                    //mylogs.WriteLine(archivo);
                    mylogs.Close();
                }

                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Open(archTemp, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                wb.SaveAs(fichero.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlXMLSpreadsheet, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                wb.Close();
                app.Quit();

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
