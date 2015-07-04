using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ExcelHelper
{
    public class ExcelExporter : IDisposable
    {
        #region Private Fields

        /// <summary>
        /// 模板Excel路径
        /// </summary>
        private string _templatePath = string.Empty;

        /// <summary>
        /// Excel应用程序类
        /// </summary>
        private Excel.Application _application;

        /// <summary>
        /// Excel工作簿
        /// </summary>
        private Excel.Workbook _workbook;

        /// <summary>
        /// Excel表单
        /// </summary>
        private Excel.Worksheet _worksheet;

        /// <summary>
        /// 当前操作的范围
        /// </summary>
        private Excel.Range _range;

        /// <summary>
        /// 当前范围字体格式
        /// </summary>
        private Excel.Font _font;

        /// <summary>
        /// 选定区域边框
        /// </summary>
        private Excel.Borders _borders;

        /// <summary>
        /// 选定区域左边框
        /// </summary>
        private Excel.Border _leftBorder;

        /// <summary>
        /// 选定区域上边框
        /// </summary>
        private Excel.Border _topBorder;

        /// <summary>
        /// 选定区域右边框
        /// </summary>
        private Excel.Border _rightBorder;

        /// <summary>
        /// 选定区域下边框
        /// </summary>
        private Excel.Border _bottomBorder;

        /// <summary>
        /// 当前工作簿索引
        /// </summary>
        private int _sheetIndex = 1;

        /// <summary>
        /// 对象是否已释放
        /// </summary>
        private bool _isDisposed = false;

        #endregion

        #region Public Properties

        /// <summary>
        /// 当前表单索引
        /// </summary>
        public int SheetIndex
        {
            get
            {
                return this._sheetIndex;
            }
            set
            {
                if(value < 1 || value > this._workbook.Sheets.Count)
                {
                    value = 1;
                }
                this._sheetIndex = value;
                this._worksheet = this._workbook.Sheets[value];
                this._range = null;
                this._font = null;
            }
        }

        /// <summary>
        /// 导出路径
        /// </summary>
        public string ExportFileName
        {
            get;
            set;
        }

        #endregion

        #region Constructors

        public ExcelExporter()
        {
            this._application = new Excel.Application();
            this._workbook = this._application.Workbooks.Add();
            this._sheetIndex = 1;
            this._worksheet = this._workbook.Sheets[this._sheetIndex];
        }

        public ExcelExporter(string templatePath)
        {
            this._templatePath = templatePath;
            this._application = new Excel.Application();
            this._workbook = this._application.Workbooks.Open(this._templatePath);
            this._sheetIndex = 1;
            this._worksheet = this._workbook.Sheets[this._sheetIndex];
        }

        ~ExcelExporter()
        {
            this.Dispose(false);
        }

        #endregion

        #region Public Methods

        #region 设置操作范围

        /// <summary>
        /// 设置单元格范围
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        public void SetRange(int row, int col)
        {
            this._range = this._worksheet.Cells[row, col];
            this._font = this._range.Font;
            this._borders = this._range.Borders;
            this._leftBorder = this._borders[Excel.XlBordersIndex.xlEdgeLeft];
            this._topBorder = this._borders[Excel.XlBordersIndex.xlEdgeTop];
            this._rightBorder = this._borders[Excel.XlBordersIndex.xlEdgeRight];
            this._bottomBorder = this._borders[Excel.XlBordersIndex.xlEdgeBottom];
        }

        /// <summary>
        /// 设置连续多单元格块范围
        /// </summary>
        /// <param name="startRow"></param>
        /// <param name="startCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        public void SetRange(int startRow, int startCol, int endRow, int endCol)
        {
            this._range = this._worksheet.Range[this._worksheet.Cells[startRow, startCol], this._worksheet.Cells[endRow, endCol]];
            this._font = this._range.Font;
            this._borders = this._range.Borders;
            this._leftBorder = this._borders[Excel.XlBordersIndex.xlEdgeLeft];
            this._topBorder = this._borders[Excel.XlBordersIndex.xlEdgeTop];
            this._rightBorder = this._borders[Excel.XlBordersIndex.xlEdgeRight];
            this._bottomBorder = this._borders[Excel.XlBordersIndex.xlEdgeBottom];
        }

        #endregion

        #region 设置单元格值

        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="value"></param>
        public void SetCellValue(int row, int col, object value)
        {
            this.SetRange(row, col);
            this._range.Value2 = value;
        }

        /// <summary>
        /// 设置合并单元格值
        /// </summary>
        /// <param name="startRow"></param>
        /// <param name="startCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        /// <param name="value"></param>
        public void SetCellValue(int startRow, int startCol, int endRow, int endCol, object value)
        {
            this.MergeCells(startRow, startCol, endRow, endCol);
            this._range.Value2 = value;
            this._range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            this._range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
        }

        #endregion

        #region 设置字体

        /// <summary>
        /// 设置当前区域格字体
        /// </summary>
        /// <param name="fontName"></param>
        public void SetFont(string fontName, int fontSize, bool isBold)
        {
            if (!string.IsNullOrWhiteSpace(fontName))
            {
                this._font.Name = fontName;
            }
            if (fontSize > -1)
            {
                this._font.Size = fontSize;
            }
            if (isBold)
            {
                this._font.Bold = isBold;
            }
        }

        /// <summary>
        /// 设置指定单元格字体
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="fontName"></param>
        public void SetFont(int row, int col, string fontName, int fontSize, bool isBold)
        {
            this.SetRange(row, col);
            this.SetFont(fontName, fontSize, isBold);
        }

        /// <summary>
        /// 设置指定范围字体样式
        /// </summary>
        /// <param name="startRow"></param>
        /// <param name="startCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        /// <param name="fontName"></param>
        /// <param name="fontSize"></param>
        /// <param name="isBold"></param>
        public void SetFont(int startRow, int startCol, int endRow, int endCol, string fontName, int fontSize, bool isBold)
        {
            this.SetRange(startRow, startCol, endRow, endCol);
            this.SetFont(fontName, fontSize, isBold);
        }

        #endregion

        #region 设置行高

        /// <summary>
        /// 设置当前范围行高
        /// </summary>
        /// <param name="height"></param>
        public void SetRowHeight(double height)
        {
            if(height > 0)
            {
                this._range.RowHeight = height;
            }
        }

        /// <summary>
        /// 设置指定行行高
        /// </summary>
        /// <param name="row"></param>
        /// <param name="height"></param>
        public void SetRowHeight(int row, double height)
        {
            this.SetRange(row, 1);
            this.SetRowHeight(height);
        }

        #endregion

        #region 设置列宽

        /// <summary>
        /// 设置当前列范围列宽
        /// </summary>
        /// <param name="width"></param>
        public void SetColumnWidth(double width)
        {
            this._range.ColumnWidth = width;
        }

        /// <summary>
        /// 设置指定列列宽
        /// </summary>
        /// <param name="col"></param>
        /// <param name="width"></param>
        public void SetColumnWidth(int col, double width)
        {
            this.SetRange(1, col);
            this.SetColumnWidth(width);
        }

        #endregion

        #region 自动调整行高列宽

        public void RowAutoFit()
        {
            this._range.EntireRow.AutoFit();
        }

        public void RowAutoFit(int row)
        {
            this.SetRange(row, 1);
            this.RowAutoFit();
        }

        public void RowAutoFit(int startRow, int endRow)
        {
            this.SetRange(startRow, 1, endRow, 1);
            this.RowAutoFit();
        }

        #endregion

        #region 自动调整列宽

        public void ColumnAutoFit()
        {
            this._range.EntireColumn.AutoFit();
        }

        public void ColumnAutoFit(int col)
        {
            this.SetRange(1, col);
            this.ColumnAutoFit();
        }

        public void ColumnAutoFit(int startCol, int endCol)
        {
            this.SetRange(1, startCol, 1, endCol);
            this.ColumnAutoFit();
        }

        #endregion

        #region 设置对齐方式

        public void SetAliment(Excel.XlHAlign horizontalAlignment, Excel.XlVAlign verticalAlignment)
        {
            this._range.HorizontalAlignment = horizontalAlignment;
            this._range.VerticalAlignment = verticalAlignment;
        }

        public void SetAliment(int row, int col, Excel.XlHAlign horizontalAlignment, Excel.XlVAlign verticalAlignment)
        {
            this.SetRange(row, col);
            this.SetAliment(horizontalAlignment, verticalAlignment);
        }

        public void SetAliment(int startRow, int startCol, int endRow, int endCol, Excel.XlHAlign horizontalAlignment, Excel.XlVAlign verticalAlignment)
        {
            this.SetRange(startRow, startCol, endRow, endCol);
            this.SetAliment(horizontalAlignment, verticalAlignment);
        }

        #endregion

        #region 设置边框

        public void SetBorders(Excel.XlLineStyle style)
        {
            this._range.Borders.LineStyle = style;
        }

        public void SetBorders(int row, int col, Excel.XlLineStyle style)
        {
            this.SetRange(row, col);
            this.SetBorders(style);
        }

        public void SetBorders(int startRow, int startCol, int endRow, int endCol, Excel.XlLineStyle style)
        {
            this.SetRange(startRow, startCol, endRow, endCol);
            this.SetBorders(style);
        }

        public void SetLeftBorder(Excel.XlLineStyle style)
        {
            this._leftBorder.LineStyle = style;
        }

        public void SetLeftBorder(int row, int col, Excel.XlLineStyle style)
        {
            this.SetRange(row, col);
            this.SetLeftBorder(style);
        }

        public void SetTopBorder(Excel.XlLineStyle style)
        {
            this._topBorder.LineStyle = style;
        }

        public void SetTopBorder(int row, int col, Excel.XlLineStyle style)
        {
            this.SetRange(row, col);
            this.SetTopBorder(style);
        }

        public void SetRightBorder(Excel.XlLineStyle style)
        {
            this._rightBorder.LineStyle = style;
        }

        public void SetRightBorder(int row, int col, Excel.XlLineStyle style)
        {
            this.SetRange(row, col);
            this.SetRightBorder(style);
        }

        public void SetBottomBorder(Excel.XlLineStyle style)
        {
            this._bottomBorder.LineStyle = style;
        }

        public void SetBottomBorder(int row, int col, Excel.XlLineStyle style)
        {
            this.SetRange(row,col);
            this.SetBottomBorder(style);
        }

        #endregion

        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="startRow"></param>
        /// <param name="startCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        public void MergeCells(int startRow, int startCol, int endRow, int endCol)
        {
            this.SetRange(startRow, startCol, endRow, endCol);
            this._range.Merge();
        }

        /// <summary>
        /// 保存
        /// </summary>
        /// <param name="fileName"></param>
        public void SaveAs(string fileName)
        {
            this._application.DisplayAlerts = false;
            this._workbook.SaveCopyAs(fileName);
            this._workbook.Close();
            this._application.Quit();
        }

        #endregion

        #region Disposable模式方法

        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool dispose)
        {
            if(this._isDisposed)
            {
                return;
            }
            if(dispose)
            {
                //此处清理托管资源，留待以后使用
            }

            //清理非托管资源
            Marshal.ReleaseComObject(this._leftBorder);
            Marshal.ReleaseComObject(this._topBorder);
            Marshal.ReleaseComObject(this._rightBorder);
            Marshal.ReleaseComObject(this._bottomBorder);
            Marshal.ReleaseComObject(this._borders);
            Marshal.ReleaseComObject(this._font);
            Marshal.ReleaseComObject(this._range);
            Marshal.ReleaseComObject(this._worksheet);
            Marshal.ReleaseComObject(this._workbook);
            Marshal.ReleaseComObject(this._application);

            this._isDisposed = true;
        }

        #endregion
    }
}
