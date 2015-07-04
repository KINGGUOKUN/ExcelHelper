using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.IO;
using ExcelHelper;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            string dir = AppDomain.CurrentDomain.BaseDirectory;
            string templatePath = string.Format(@"{0}{1}", dir, "账目出入.xlsx");

            #region 导入测试
            ExcelImporter importer = new ExcelImporter();
            DataTable dt = null;
            try
            {
                dt = importer.Import(templatePath, "Sheet 1");
            }
            catch(Exception e)
            {
                return;
            }
            List<Entity> list = new List<Entity>();
            Entity entity;
            foreach(DataRow dr in dt.Rows)
            {
                entity = new Entity()
                {
                    Code = dr.Field<string>("单位编号"),
                    Name = dr.Field<string>("单位名称"),
                };
                string account = dr.Field<string>("应付余额");
                decimal Account;
                decimal.TryParse(account, out Account);
                entity.Account = Account;
                list.Add(entity);
            }
            
            #endregion

            #region 导出测试
            //ExcelUtil util = new ExcelUtil(templatePath);
            //using (ExcelUtil util = new ExcelUtil(templatePath))
            //{
            //    //util.SetCellValue(2, 1, "0810311211");
            //    //util.SetCellValue(2, 2, "郭坤");
            //    //util.SetFont("楷体_GB2312", 20, true);
            //    //util.SetCellValue(2, 3, "男");
            //    //util.SetCellValue(2, 4, 25);
            //    //util.RowAutoFit(2, 4);
            //    //util.ColumnAutoFit(1, 4);

            //    for (int i = 2; i <= 10;i++ )
            //    {
            //        util.SetCellValue(i, 1, "学号" + (i - 1).ToString());
            //        util.SetCellValue(i, 2, "姓名"  + (i - 1).ToString());
            //        util.SetCellValue(i, 3, "性别" + (i - 1).ToString());
            //        util.SetCellValue(i, 4, "年龄" + (i - 1).ToString());
            //    }
            //    util.SetRange(2, 1, 10, 4);
            //    util.SetAliment(Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignCenter);
            //    //util.SetBorders(Excel.XlLineStyle.xlSlantDashDot);
            //    util.SetLeftBorder(Excel.XlLineStyle.xlContinuous);
            //    string exportPath = string.Format(@"{0}{1}", dir, "1.xlsx");
            //    util.SaveAs(exportPath);
            //}  

            //using (ExcelExporter util = new ExcelExporter())
            //{
            //    //util.SetCellValue(2, 1, "0810311211");
            //    //util.SetCellValue(2, 2, "郭坤");
            //    //util.SetFont("楷体_GB2312", 20, true);
            //    //util.SetCellValue(2, 3, "男");
            //    //util.SetCellValue(2, 4, 25);
            //    //util.RowAutoFit(2, 4);
            //    //util.ColumnAutoFit(1, 4);

            //    for (int i = 1; i <= 10; i++)
            //    {
            //        util.SetCellValue(i, 1, "学号" + i.ToString());
            //        util.SetCellValue(i, 2, "姓名" + i.ToString());
            //        util.SetCellValue(i, 3, "性别" + i.ToString());
            //        util.SetCellValue(i, 4, "年龄" + i.ToString());
            //    }
            //    util.SetRange(1, 1, 10, 4);
            //    util.SetAliment(Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignCenter);
            //    //util.SetBorders(Excel.XlLineStyle.xlSlantDashDot);
            //    util.SetLeftBorder(Excel.XlLineStyle.xlContinuous);
            //    string exportPath = string.Format(@"{0}{1}", dir, "1.xlsx");
            //    util.SaveAs(exportPath);
            //}  
            #endregion

            Console.Read();
  
        }
    }

    internal class Entity
    {
        public string Code { get; set; }

        public string Name { get; set; }

        public decimal Account { get; set; }
    }
}
