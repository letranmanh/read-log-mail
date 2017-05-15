using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
namespace ReadFileToExcel
{
    class ExcelReader
    {
        private Microsoft.Office.Interop.Excel.Application app = null;
        private Microsoft.Office.Interop.Excel.Workbook workbook = null;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
        private Microsoft.Office.Interop.Excel.Range workSheet_range = null;

        public void ExportExcel()
        {
            try
            {
                app = new Microsoft.Office.Interop.Excel.Application();
                app.Visible = false;

                workbook = app.Workbooks.Add(1);
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];


                //open the file to write
            }
            catch (Exception e)
            {
                Console.Write(e.ToString());
            }
            finally
            {
            }
        }
        public void addData(int row, int col, string data)
        {
            worksheet.Cells[row, col] = data;  //writes data
        }

        public void close(string s)
        {
            String path = s;
            workbook.SaveAs(path + @"\filename", Excel.XlFileFormat.xlWorkbookNormal, null, null, null, null, Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);

            //saved it to the path


            workbook.Close(true, null, null);


            //and finally closed the file.

        }
    }
}
