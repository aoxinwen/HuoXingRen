using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MSExel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    class Class1_excel
    {
        public static MSExel.Application excel { get; set; }
        public static MSExel.Workbooks wbs { get; set; }
        public static MSExel._Workbook wb { get; set; }
        public static MSExel._Worksheet exsheet { get; set; }
        public static MSExel.Range DYG { get; set; }

        /// <summary>
        /// 打开excel表格获取EXCEL文件wb
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static MSExel._Workbook Open_excel(string fileName)
        {
            excel = new MSExel.Application();
            /*{ 
                Visible = true//设置文件打开后可见
            }; */          
            wbs = excel.Workbooks;
            wb = wbs.Add(fileName);
            return wb;
        }

        /// <summary>
        /// 关闭文件和释放资源
        /// </summary>
        /// <returns></returns>
        public static bool Close_excel()
        {
            try
            {
                wb.Close();//关闭文件
                wbs.Close();//关闭工作簿
                excel.Quit();//关闭excel进程
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);//释放excel进程资源
                return true;
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message, "提示", MessageBoxButtons.OK);
                return false;
            }
        }

        /// <summary>
        /// 获取一个sheet
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public static bool Get_sheet(int index)
        {
            try
            {
                exsheet = (MSExel.Worksheet)wb.Sheets[index];
                exsheet.Activate();
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 删除sheet
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public static bool Del_sheet(int index)
        {
            try
            {
                exsheet = (MSExel.Worksheet)wb.Sheets.get_Item(index);
                excel.DisplayAlerts = false;
                exsheet.Delete();
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 获取单元格
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        public static bool Get_cell(int x, int y)
        {
            try
            {
                DYG = (MSExel.Range)exsheet.Cells[x, y];
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
