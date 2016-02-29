using System;
using System.Windows;

namespace ExcelToolsGlobalNamingArea
{
    internal class ExcelCreation : IDisposable
    {
        private ExcelToolsVM vm;
        //private Microsoft.Office.Interop.Excel.Application xlApp = null;
        //private WorkBooks xlBooks = null;
        //private WorkBook xlBook = null;
        //private Worksheets xlSheets = null;
        //private Worksheet xlSheet = null;

        private enum ReleaseMode
        {
            Sheet,
            Sheets,
            Book,
            Books,
            App
        }

        public ExcelCreation(ExcelToolsVM vm)
        {
            this.vm = vm;

            // create ExcelBook
            // xlApp = new Microsoft.Office.Interop.Excel.Application();

            // Add sheet
            //xlBook = xlApp.WorkBooks.Add();

            // create sheets
            this.vm.MaxLoopCnt += 2;

            for (int i = 0; i < this.vm.MaxLoopCnt; i++)
            {
                // add sheet if the sheet is more than 4
                if (i < 4)
                {
                    continue;
                }

                // add sheet
                //xlSheet = xlBook.Sheets.Add();
            }

            // change sheet name
            for (int i = 1; i < this.vm.MaxLoopCnt; i++)
            {
                // get the sheet
                // xlSheet = xlBook.Sheets[i] as Microsoft.Office.Interop.Excel.Worksheet;

                if (!string.IsNullOrEmpty(this.vm.SheetNm))
                {
                    string harmlessNm = this.vm.SheetNm
                        .Replace(":", string.Empty)
                        .Replace("\\", string.Empty)
                        .Replace("/", string.Empty)
                        .Replace("?", string.Empty)
                        .Replace("*", string.Empty)
                        .Replace("[", string.Empty)
                        .Replace("]", string.Empty)
                        .Replace("'", string.Empty)
                        .Replace("：", string.Empty)
                        .Replace("￥", string.Empty)
                        .Replace("／", string.Empty)
                        .Replace("？", string.Empty)
                        .Replace("＊", string.Empty)
                        .Replace("［", string.Empty)
                        .Replace("］", string.Empty)
                        .Replace("’", string.Empty);

                    //xlSheet.Name = harmlessNm.Trim() + (this.vm.No1 - 1 + i).ToString();
                }
                else
                {
                    //xlSheet.Name = (this.vm.No1 - 1 + i).ToString();
                }
            }

            System.Threading.Thread.Sleep(1000);

            // Show Excel Book
            // xlApp.Visible = true;
        }

        void IDisposable.Dispose()
        {
            ReleaseExcelComObject(ReleaseMode.App);

            GC.SuppressFinalize(this);
        }

        private void ReleaseExcelComObject(ReleaseMode app)
        {
            try
            {
                // Release resources
            }
            catch (Exception)
            {
                MessageBox.Show("Falied release resources. Launch the task manager and kill the EXCEL.EXE process, please."
                    , "Error"
                    , MessageBoxButton.OK
                    , MessageBoxImage.Error);
                throw;
            }
        }
    }
}