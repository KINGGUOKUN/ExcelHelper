using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace ExcelHelper
{
    public class ExcelImporter
    {
        #region Private Fields

        private readonly string strConnection1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR=NO;IMEX=1\"";
        private readonly string strConnection2 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0;HDR=NO;IMEX=1\"";

        #endregion

        #region Constructors

        public ExcelImporter()
        {
        }

        #endregion

        #region Public Methods

        public DataTable Import(string filePath, string sheetName)
        {
            DataTable result = new DataTable();
            FileInfo info = new FileInfo(filePath);
            string extension = info.Extension;
            string connectionString = string.Empty;
            if (extension == ".xls")
            {
                connectionString = string.Format(strConnection1, info.FullName);
            }
            else if (extension == ".xlsx")
            {
                connectionString = string.Format(strConnection2, info.FullName);
            }
            else
            {
                throw new Exception("指定路径文件非Excel文件");
            }
            using(OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();
                }
                catch(OleDbException e)
                {
                    throw new Exception("打开指定Excel文件出错", e);
                }
                string strCmd = string.Format("select * from [{0}$]", sheetName);
                using (OleDbCommand cmd = new OleDbCommand(strCmd, connection))
                {
                    OleDbDataAdapter adapter = new OleDbDataAdapter(cmd);
                    try
                    {
                        adapter.Fill(result);
                    }
                    catch(OleDbException e)
                    {
                        throw new Exception("读取指定工作簿出错", e);
                    } 
                }
            }
            DataRow drColumns = result.Rows[0];
            for (int i = 0; i < result.Columns.Count; i++)
            {
                result.Columns[i].ColumnName = drColumns[i].ToString();
            }
            drColumns.Delete();
            result.TableName = sheetName;
            result.AcceptChanges();
            return result;
        }

        #endregion
    }
}
