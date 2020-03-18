using ExcelProject.Entities;
using System;
using System.Collections.Generic;

namespace ExcelProject.DataAccess.Concrete.Context
{
    public class Excel
    {
        /// Kaynak yolu verilen dosyanın içeriği okunur.       
        public List<T_GK_MAIN> Read(string filePath)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wb = excelApp.Workbooks.Open(filePath);
            Microsoft.Office.Interop.Excel.Worksheet currentPage = null;
            List<T_GK_MAIN> tgkMain = new List<T_GK_MAIN>();
            for (int i = 1; i <= 2; i++)
            {
                currentPage = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(i);
                tgkMain.Add(FillModel(currentPage, i));
            }
            wb.Close();
            excelApp.Quit();
            ClearRam(excelApp);
            ClearRam(wb);
            ClearRam(currentPage);
            return tgkMain;
        }

        /// Listeyi T_GK modeline çevirir.       
        public List<T_GK> FillItem(dynamic list)
        {
            if (list == null)
                return null;
            List<T_GK> items = new List<T_GK>();
            int ndx = 0;
            foreach (var item in list)
            {
                ndx++;
                if (ndx < 3)
                    continue;
                if (item == null)
                    break;
                if (ndx == 2 && !(Convert.ToString(item)).Contains("T_GK"))
                    break;
                items.Add(new T_GK() { code = Convert.ToString(item) });
            }
            return items;
        }

        /// Sayfadaki sütunları T_GK_MAIN modeline çevirir.       
        public T_GK_MAIN FillModel(Microsoft.Office.Interop.Excel.Worksheet page, int pageId)
        {
            T_GK_MAIN tgkMain = new T_GK_MAIN();
            tgkMain.id = pageId;
            Microsoft.Office.Interop.Excel.Range rng = null;
            rng = (Microsoft.Office.Interop.Excel.Range)page.Columns[1];
            tgkMain.tgk1 = FillItem(rng.Value2);
            rng = (Microsoft.Office.Interop.Excel.Range)page.Columns[2];
            tgkMain.tgk2 = FillItem(rng.Value2);
            rng = (Microsoft.Office.Interop.Excel.Range)page.Columns[3];
            tgkMain.tgk3 = FillItem(rng.Value2);
            rng = (Microsoft.Office.Interop.Excel.Range)page.Columns[4];
            tgkMain.tgk4 = FillItem(rng.Value2);
            rng = (Microsoft.Office.Interop.Excel.Range)page.Columns[5];
            tgkMain.tgk5 = FillItem(rng.Value2);
            return tgkMain;
        }

        /// Ram temizleme metodu       
        public void ClearRam(object app)
        {
            try
            {
                if (app == null)
                    return;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            }
            catch (Exception)
            {
            }
            finally
            {
                GC.Collect();
            }
        }
    }

}
