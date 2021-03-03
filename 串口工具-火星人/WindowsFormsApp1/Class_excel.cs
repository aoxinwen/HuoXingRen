using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    class Class_excel
    {
        public static IWorkbook wb { get; set; }         //工作簿
        public static ISheet exsheet { get; set; }       //工作表
        public static IRow row { get; set; }             //某行
        public static ICell cell { get; set; }           //某单元格
        


        /// <summary>
        /// 打开excel表格获取EXCEL文件wb
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static bool Open_excel(string fileName)
        {
            string extension = Path.GetExtension(fileName);
            try
            {
                using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    if (extension.Equals(".xls"))
                    {
                        wb = new HSSFWorkbook(fs);
                    }
                    else
                    {
                        wb = new XSSFWorkbook(fs);

                    }
                    fs.Close();
                }                                
                return true;
            }
            catch
            {
                MessageBox.Show("请检查路径", "提示", MessageBoxButtons.OK);
                return false;
            }
        }

        /// <summary>
        /// 关闭文件和释放资源
        /// </summary>
        /// <returns></returns>
        public static bool Close_excel()
        {
            try
            {
                wb.Close();//关闭工作簿
                return true;
            }
            catch (Exception e)
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
                exsheet = wb.GetSheetAt(index);          
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
                wb.RemoveSheetAt(index);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static bool Get_row(int index)
        {
            try
            {
                row = exsheet.GetRow(index);
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
                cell = exsheet.GetRow(x).GetCell(y, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 写入文件
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static bool Write_sheet(string fileName)
        {
            try
            {
                using (FileStream fs = File.Open(fileName, FileMode.OpenOrCreate))
                {
                    wb.Write(fs);
                    fs.Close();
                }
                    return true;
            }
            catch
            {
                return false;
            }
        }
    }

}
