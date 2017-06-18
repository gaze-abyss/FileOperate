using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Data;
using Xls = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using System.Collections.ObjectModel;
using System.Data.OleDb;
using Microsoft.Win32;

namespace FileOperate.Excel
{
    /// <summary>
    /// Excel 处理类
    /// </summary>
    /// <remarks> 可以用于创建 Excel ，操作工作表，设置单元格样式对齐方式等，导入内存、数据库中的数据表，插入图片到 Excel 等</remarks>
    public sealed class ExcelHelper : IDisposable
    {
        #region 构造函数
        /// <summary>
        /// ExcelHandler 的构造函数
        /// </summary>
        /// <param name="fileName"> Excel 文件名，绝对路径 </param>
        public ExcelHelper(string fileName)
            : this(fileName, false)
        {
        }
        /// <summary>
        /// 创建 ExcelHandler 对象，指定文件名以及是否创建新的 Excel 文件
        /// </summary>
        /// <param name="fileName"> Excel 文件名，绝对路径 </param>
        /// <param name="createNew"> 是否创建新的 Excel 文件 </param>
        public ExcelHelper(string fileName, bool createNew)
        {
            this.FileName = fileName;
            this.ifCreateNew = createNew;
        }
        #endregion

        #region 字段和属性
        private static readonly object _missing = Missing.Value;
        private string _fileName;
        /// <summary>
        /// Excel 文件名
        /// </summary>
        public string FileName
        {
            get { return _fileName; }
            set { _fileName = value; }
        }
        /// <summary>
        /// 是否新建 Excel 文件
        /// </summary>
        private bool ifCreateNew;
        private Xls.Application _app;
        /// <summary>
        /// 当前 Excel 应用程序
        /// </summary>
        public Xls.Application App
        {
            get { return _app; }
            set { _app = value; }
        }
        private Xls.Workbooks _allWorkbooks;
        /// <summary>
        /// 当前 Excel 应用程序所打开的所有 Excel 工作簿
        /// </summary>
        public Xls.Workbooks AllWorkbooks
        {
            get { return _allWorkbooks; }
            set { _allWorkbooks = value; }
        }
        private Xls.Workbook _currentWorkbook;
        /// <summary>
        /// 当前 Excel 工作簿
        /// </summary>
        public Xls.Workbook CurrentWorkbook
        {
            get { return _currentWorkbook; }
            set { _currentWorkbook = value; }
        }
        private Xls.Worksheets _allWorksheets;
        /// <summary>
        /// 当前 Excel 工作簿内的所有 Sheet
        /// </summary>
        public Xls.Worksheets AllWorksheets
        {
            get { return _allWorksheets; }
            set { _allWorksheets = value; }
        }
        private Xls.Worksheet _currentWorksheet;
        /// <summary>
        /// 当前 Excel 中激活的 Sheet
        /// </summary>
        public Xls.Worksheet CurrentWorksheet
        {
            get { return _currentWorksheet; }
            set { _currentWorksheet = value; }
        }
        #endregion

        #region 初始化操作，打开或者创建文件
        /// <summary>
        /// 初始化，如果不创建新文件直接打开，否则创建新文件
        /// </summary>
        public bool OpenOrCreate()
        {
            try
            {
                //if (this.ExistsRegedit() == 0)
                //{
                //    //MessageBox.Show("无法创建Excel对象，可能您的计算机未正确安装Excel!");
                //    return false;
                //}
                this.App = new Xls.Application();
                if (this.App == null)
                {
                    //MessageBox.Show("无法创建Excel对象，可能您的计算机未正确安装Excel!");
                    return false;
                }
                this.AllWorkbooks = this.App.Workbooks;
                if (!this.ifCreateNew) // 直接打开
                {
                    if (!File.Exists(this.FileName))
                    {
                        throw new FileNotFoundException("找不到指定的 Excel 文件，请检查路径是否正确！ ", this.FileName);
                    }
                    this.CurrentWorkbook = this.AllWorkbooks.Open(this.FileName,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Xls.XlPlatform.xlWindows, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
                else // 创建新文件
                {
                    if (File.Exists(this.FileName))
                    {
                        File.Delete(this.FileName);
                    }
                    this.CurrentWorkbook = this.AllWorkbooks.Add(Type.Missing);
                }
                this.AllWorksheets = this.CurrentWorkbook.Worksheets as
                Xls.Worksheets;
                this.CurrentWorksheet = this.CurrentWorkbook.ActiveSheet as
                Xls.Worksheet;
                this.App.DisplayAlerts = false;
                this.App.Visible = false;
                return true;
            }
            catch (Exception x)
            {
                //try
                //{
                //    var typeCLSID = Guid.Parse("000208D5-0000-0000-C000-000000000046");
                //    Type type = System.Type.GetTypeFromCLSID(typeCLSID);
                //    object excelComObj = Activator.CreateInstance(type);
                //}
                //catch (Exception ex)
                //{
                //    MessageBox.Show("无法创建Excel对象，可能您的计算机未安装Excel!");
                //    return false;
                //}
                //MessageBox.Show("无法创建Excel对象，可能您的计算机未正确安装Excel!");
                return false;
            }

        }
        #endregion

        #region Excel Sheet 相关操作等
        public object[,] GetSheetAllData(Xls.Worksheet sheet,out int row,out int column)
        {
            column = sheet.Range["IV1"].End[Xls.XlDirection.xlToLeft].Column;
            row = sheet.Range["A65535"].End[Xls.XlDirection.xlUp].Row;
            Xls.Range rang = GetRange(sheet, 1, 1, row, column);
            object[,] rangStr = new object[2, 2];
            rangStr = rang.Value2;
            return rangStr;
        }
        /// <summary>
        /// 根据工作表名获取 Excel 工作表对象的引用
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public Xls.Worksheet GetWorksheet(string sheetName)
        {
            return this.CurrentWorkbook.Sheets[sheetName] as Xls.Worksheet;
        }
        /// <summary>
        /// 根据工作表索引获取 Excel 工作表对象的引用
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public Xls.Worksheet GetWorksheet(int index)
        {
            return this.CurrentWorkbook.Sheets.get_Item(index) as Xls.Worksheet;
        }
        /// <summary>
        /// 给当前工作簿添加工作表并返回的方法重载 ,添加工作表后不使其激活
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public Xls.Worksheet AddWorksheet(string sheetName)
        {
            return this.AddWorksheet(sheetName, false);
        }
        /// <summary>
        /// 给当前工作簿添加工作表并返回
        /// </summary>
        /// <param name="sheetName"> 工作表名 </param>
        /// <param name="activated"> 创建后是否使其激活 </param>
        /// <returns></returns>
        public Xls.Worksheet AddWorksheet(string sheetName, bool activated)
        {
            Xls.Worksheet sheet =
            this.CurrentWorkbook.Worksheets.Add(Type.Missing, Type.Missing, 1,
            Type.Missing) as Xls.Worksheet;
            sheet.Name = sheetName;
            if (activated)
            {
                sheet.Activate();
            }
            return sheet;
        }
        /// <summary>
        /// 重命名工作表
        /// </summary>
        /// <param name="sheet"> 工作表对象 </param>
        /// <param name="newName"> 工作表新名称 </param>
        /// <returns></returns>
        public Xls.Worksheet RenameWorksheet(Xls.Worksheet sheet, string
        newName)
        {
            sheet.Name = newName;
            return sheet;
        }
        /// <summary>
        /// 重命名工作表
        /// </summary>
        /// <param name="oldName"> 原名称 </param>
        /// <param name="newName"> 新名称 </param>
        /// <returns></returns>
        public Xls.Worksheet RenameWorksheet(string oldName, string newName)
        {
            Xls.Worksheet sheet = this.GetWorksheet(oldName);
            return this.RenameWorksheet(sheet, newName);
        }
        /// <summary>
        /// 删除工作表
        /// </summary>
        /// <param name="sheetName"> 工作表名 </param>
        public void DeleteWorksheet(string sheetName)
        {
            if (this.CurrentWorkbook.Worksheets.Count <= 1)
            {
                throw new InvalidOperationException("工作簿至少需要一个可视化的工作表！ ");
            }
            this.GetWorksheet(sheetName).Delete();
        }
        /// <summary>
        /// 删除除参数 sheet 指定外的其余工作表
        /// </summary>
        /// <param name="sheet"></param>
        public void DeleteWorksheetExcept(Xls.Worksheet sheet)
        {
            foreach (Xls.Worksheet ws in this.CurrentWorkbook.Worksheets)
            {
                if (sheet != ws)
                {
                    ws.Delete();
                }
            }
        }
        #endregion

        #region 单元格， Range 相关操作
        /// <summary>
        /// 设置单元格的值
        /// </summary>
        /// <param name="sheet"> 工作表 </param>
        /// <param name="rowNumber"> 单元格行号 </param>
        /// <param name="columnNumber"> 单元格列号 </param>
        /// <param name="value"> 单元格值 </param>
        public void SetCellValue(Xls.Worksheet sheet, int rowNumber, int
        columnNumber, object value)
        {
            sheet.Cells[rowNumber, columnNumber] = value;
        }
        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="sheet"> 工作表 </param>
        /// <param name="rowNumber1"> 第一个单元格行号 </param>
        /// <param name="columnNumber1"> 第一个单元格列号 </param>
        /// <param name="rowNumber2"> 结束单元格行号 </param>
        /// <param name="columnNumber2"> 结束单元格列号 </param>
        public void MergeCells(Xls.Worksheet sheet, int rowNumber1, int
        columnNumber1, int rowNumber2, int columnNumber2)
        {
            Xls.Range range = this.GetRange(sheet, rowNumber1,
            columnNumber1, rowNumber2, columnNumber2);
            range.Merge(Type.Missing);
        }
        /// <summary>
        /// 获取 Range 对象
        /// </summary>
        /// <param name="sheet"> 工作表 </param>
        /// <param name="rowNumber1"> 第一个单元格行号 </param>
        /// <param name="columnNumber1"> 第一个单元格列号 </param>
        /// <param name="rowNumber2"> 结束单元格行号 </param>
        /// <param name="columnNumber2"> 结束单元格列号 </param>
        /// <returns></returns>
        public Xls.Range GetRange(Xls.Worksheet sheet, int rowNumber1, int
        columnNumber1, int rowNumber2, int columnNumber2)
        {
            return sheet.Range[sheet.Cells[rowNumber1, columnNumber1],
            sheet.Cells[rowNumber2, columnNumber2]];
        }
        #endregion

        #region 设置单元格、 Range 的样式、对齐方式自动换行等
        /// <summary>
        /// 自动调整，设置自动换行以及自动调整列宽
        /// </summary>
        /// <param name="range"></param>
        public void AutoAdjustment(Xls.Range range)
        {
            range.WrapText = true;
            range.AutoFit();
        }
        /// <summary>
        /// 设置 Range 的单元格样式
        /// </summary>
        /// <remarks> 将各项值设置为默认值 </remarks>
        /// <param name="range"></param>
        public void SetRangeFormat(Xls.Range range)
        {
            this.SetRangeFormat(range, 11, Xls.Constants.xlAutomatic,
            Xls.Constants.xlColor1, Xls.Constants.xlLeft);
        }
        /// <summary>
        /// 设置 Range 的单元格样式
        /// </summary>
        /// <remarks> 将各项值设置为默认值 </remarks>
        /// <param name="sheet"></param>
        /// <param name="rowNumber1"></param>
        /// <param name="columnNumber1"></param>
        /// <param name="rowNumber2"></param>
        /// <param name="columNumber2"></param>
        public void SetRangeFormat(Xls.Worksheet sheet, int rowNumber1, int
        columnNumber1, int rowNumber2, int columNumber2)
        {
            this.SetRangeFormat(sheet, rowNumber1, columnNumber1,
            rowNumber2, columNumber2, 11, Xls.Constants.xlAutomatic);
        }
        /// <summary>
        /// 设置 Range 的单元格样式
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowNumber1"> 第一个单元格行号 </param>
        /// <param name="columnNumber1"> 第一个单元格列号 </param>
        /// <param name="rowNumber2"> 结束单元格行号 </param>
        /// <param name="columnNumber2"> 结束单元格列号 </param>
        /// <param name="fontSize"></param>
        /// <param name="fontName"></param>
        public void SetRangeFormat(Xls.Worksheet sheet, int rowNumber1, int
        columnNumber1, int rowNumber2, int columNumber2, object fontSize, object
        fontName)
        {
            this.SetRangeFormat(this.GetRange(sheet, rowNumber1,
            columnNumber1, rowNumber2, columNumber2), fontSize, fontName,
            Xls.Constants.xlColor1, Xls.Constants.xlLeft);
        }
        /// <summary>
        /// 设置 Range 的单元格样式
        /// </summary>
        /// <param name="range"> Range 对象 </param>
        /// <param name="fontSize"> 字体大小 </param>
        /// <param name="fontName"> 字体名称 </param>
        /// <param name="color"> 字体颜色 </param>
        /// <param name="horizontalAlignment"> 水平对齐方式 </param>
        public void SetRangeFormat(Xls.Range range, object fontSize, object
        fontName, Xls.Constants color, Xls.Constants horizontalAlignment)
        {
            range.Font.Color = color;
            range.Font.Size = fontSize;
            range.Font.Name = fontName;
            range.HorizontalAlignment = horizontalAlignment;
        }

        #endregion

        #region 导入内存中的 DataTable
        /// <summary>
        /// 导入内存中的数据表到 Excel 中
        /// </summary>
        /// <remarks> 直接导入到工作表的最起始部分 </remarks>
        /// <param name="sheet"></param>
        /// <param name="headerTitle"></param>
        /// <param name="showTitle"></param>
        /// <param name="headers"></param>
        /// <param name="table"></param>
        public void ImportDataTable(Xls.Worksheet sheet, string headerTitle, bool
        showTitle, object[] headers, DataTable table)
        {
            this.ImportDataTable(sheet, headerTitle, showTitle, headers, 1, 1, table);
        }

        /// <summary>
        /// 导入内存中的数据表到 Excel 中
        /// </summary>
        /// <remarks> 直接导入到工作表的最起始部分，且不显示标题行 </remarks>
        /// <param name="sheet"></param>
        /// <param name="headers"></param>
        /// <param name="table"></param>
        public void ImportDataTable(Xls.Worksheet sheet, object[] headers, DataTable table)
        {
            this.ImportDataTable(sheet, null, false, headers, table);
        }
        /// <summary>
        /// 导入内存中的数据表到 Excel 中
        /// </summary>
        /// <remarks> 标题行每一列与 DataTable 标题一致 </remarks>
        /// <param name="sheet"></param>
        /// <param name="table"></param>
        public void ImportDataTable(Xls.Worksheet sheet, DataTable table)
        {
            List<string> headers = new List<string>();
            foreach (DataColumn column in table.Columns)
            {
                headers.Add(column.Caption);
            }
            this.ImportDataTable(sheet, headers.ToArray(), table);
        }
        /// <summary>
        /// 导入内存中的数据表到 Excel 中
        /// </summary>
        /// <param name="sheet"> 工作表 </param>
        /// <param name="headerTitle"> 表格标题 </param>
        /// <param name="showTitle"> 是否显示表格标题行 </param>
        /// <param name="headers"> 表格每一列的标题 </param>
        /// <param name="rowNumber"> 插入表格的起始行号 </param>
        /// <param name="columnNumber"> 插入表格的起始列号 </param>
        /// <param name="table"> 内存中的数据表 </param>
        public void ImportDataTable(Xls.Worksheet sheet, string headerTitle, bool showTitle, object[] headers, int rowNumber, int columnNumber, DataTable table)
        {
            int columns = table.Columns.Count;
            int rows = table.Rows.Count;
            int titleRowIndex = rowNumber;
            int headerRowIndex = rowNumber;
            Xls.Range titleRange = null;
            if (showTitle)
            {
                headerRowIndex++;
                // 添加标题行，并设置样式
                titleRange = this.GetRange(sheet, rowNumber, columnNumber,
                rowNumber, columnNumber + columns - 1);
                titleRange.Merge(_missing);
                this.SetRangeFormat(titleRange, 16, Xls.Constants.xlAutomatic,
                Xls.Constants.xlColor1, Xls.Constants.xlCenter);
                titleRange.Value2 = headerTitle;
            }
            // 添加表头
            int m = 0;
            foreach (object header in headers)
            {
                this.SetCellValue(sheet, headerRowIndex, columnNumber + m, header);
                m++;
            }
            // 添加每一行的数据
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    sheet.Cells[headerRowIndex + i + 1, j + columnNumber] =
                    table.Rows[i][j];
                }
            }
        }
        #endregion

        #region 插入图片到 Excel 中的相关方法
        /// <summary>
        /// 插入图片
        /// </summary>
        /// <param name="sheet"> 工作表 </param>
        /// <param name="imageFilePath"> 图片的绝对路径 </param>
        /// <param name="rowNumber"> 单元格行号 </param>
        /// <param name="columnNumber"> 单元格列号 </param>
        /// <returns></returns>
        public Xls.Picture AddImage(Xls.Worksheet sheet, string imageFilePath, int
        rowNumber, int columnNumber)
        {
            Xls.Range range = this.GetRange(sheet, rowNumber, columnNumber,
            rowNumber, columnNumber);
            range.Select();
            Xls.Pictures pics = sheet.Pictures(_missing) as Xls.Pictures;
            Xls.Picture pic = pics.Insert(imageFilePath, _missing);
            pic.Left = (double)range.Left;
            pic.Top = (double)range.Top;
            return pic;
        }
        /// <summary>
        /// 插入图片
        /// </summary>
        /// <param name="sheet"> 工作表 </param>
        /// <param name="imageFilePath"> 图片的绝对路径 </param>
        /// <param name="rowNumber"> 单元格行号 </param>
        /// <param name="columnNumber"> 单元格列号 </param>
        /// <param name="width"> 图片的宽度 </param>
        /// <param name="height"> 图片的高度 </param>
        /// <returns></returns>
        public Xls.Picture AddImage(Xls.Worksheet sheet, string imageFilePath, int
        rowNumber, int columnNumber, double width, double height)
        {
            Xls.Picture pic = this.AddImage(sheet, imageFilePath, rowNumber,
            columnNumber);
            pic.Width = width;
            pic.Height = height;
            return pic;
        }
        ///// <summary>
        /////  插入图片
        ///// </summary>
        ///// <remarks> 从流中读取图片 </remarks>
        ///// <param name="sheet"></param>
        ///// <param name="imageStream"></param>
        ///// <param name="x"></param>
        ///// <param name="y"></param>
        ///// <param name="width"></param>
        ///// <param name="height"></param>
        ///// <returns></returns>
        //public Xls.Picture AddImage(Xls.Worksheet sheet, Stream imageStream,
        //int x, int y, double width, double height)
        //{
        //}
        #endregion

        #region 保存 Excel
        /// <summary>
        /// 保存 Excel
        /// </summary>
        public void Save()
        {
            if (this.ifCreateNew)
            {
                this.SaveAs(this.FileName);
            }
            else
            {
                this.CurrentWorkbook.Save();
            }
            //this.SaveAs(this.FileName);
        }
        /// <summary>
        /// 保存 Excel
        /// </summary>
        /// <param name="filePath"> 文件的绝对路径 </param>
        public void SaveAs(string filePath)
        {
            this.CurrentWorkbook.SaveAs(filePath,
            Xls.XlFileFormat.xlWorkbookNormal, _missing, _missing, _missing,
            _missing, Xls.XlSaveAsAccessMode.xlNoChange, _missing, _missing,
            _missing, _missing, _missing);
        }
        #endregion

        #region IDisposable 成员
        /// <summary>
        /// 对象销毁时执行的操作
        /// </summary>
        public void Dispose()
        {
            this.CurrentWorkbook.Close(true, this.FileName, _missing);
            Marshal.FinalReleaseComObject(this.CurrentWorkbook);
            this.CurrentWorkbook = null;
            this.App.Quit();
            Marshal.FinalReleaseComObject(this.App);
            this.App = null;
            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
        }
        #endregion

        #region 查询注册表，判断本机是否安装Office2003,2007和WPS
        /// <summary>
        /// 查询注册表，判断本机是否安装Office2003,2007和WPS
        /// </summary>
        /// <returns></returns>
        public int ExistsRegedit()
        {
            int ifused = 0;
            RegistryKey rk = Registry.LocalMachine;
            //查询Office2003
            RegistryKey f03 = rk.OpenSubKey(@"SOFTWARE\Microsoft\Office\11.0\Excel\InstallRoot\");
            //查询Office2007
            RegistryKey f07 = rk.OpenSubKey(@"SOFTWARE\Microsoft\Office\12.0\Excel\InstallRoot\");
            //查询wps
            //RegistryKey wps = rk.OpenSubKey(@"Software\Kingsoft\Office\6.0\common");
            RegistryKey wps = Registry.CurrentUser.OpenSubKey(@"Software\Kingsoft\Office\6.0\common");

            //检查本机是否安装Office2003
            if (f03 != null)
            {
                string file03 = f03.GetValue("Path").ToString();
                if (File.Exists(file03 + "Excel.exe")) ifused += 1;
            }

            //检查本机是否安装Office2007
            if (f07 != null)
            {
                string file07 = f07.GetValue("Path").ToString();
                if (File.Exists(file07 + "Excel.exe")) ifused += 2;
            }

            //检查本机是否安装wps
            if (wps != null)
            {
                if (wps.GetValue("InstallRoot") != null)
                {
                    string strpath = wps.GetValue("InstallRoot").ToString();
                    if (File.Exists(strpath + @"\office6\wps.exe"))
                    {
                        ifused += 4;
                    }
                }
            }

            return ifused;
        }
        #endregion

        #region 测试函数
        /// <summary>
        /// 表格文件测试函数
        /// </summary>
        public void TestFile()
        {
            ///按时间创建表格
            ///string  excelFilePath  = string.Format("{0}Excel-{1}.xls",  AppDomain.CurrentDomain.BaseDirectory,  DateTime.Now.ToString("yyyyMMddHHmmss"));
            string excelFilePath = string.Format("{0}Excel_Test.xls", AppDomain.CurrentDomain.BaseDirectory);
            ExcelHelper handler;
            ///创建Excel
            //ExcelHelper  handler  = new ExcelHelper(excelFilePath,  true);
            {
                //handler.OpenOrCreate();
                //MessageBox.Show(" 创建 Excel 成功！ ");
                //handler.Save();
                //MessageBox.Show(string.Format(" 保存 Excel 成功！ Excel 路径 :{0}", excelFilePath));
                //handler.Dispose();
            }
            //创建一个工作表 设置第二个参数为 false 表示直接打开现有的 Excel 文档
            {
                //handler = new ExcelHelper(excelFilePath, false);
                //handler.OpenOrCreate();
                //// 创建一个 Worksheet
                //Worksheet sheet = handler.AddWorksheet("TestSheet");
                //// 删除除 TestSheet 之外的其余 Worksheet
                //handler.DeleteWorksheetExcept(sheet);
                //handler.Save();
                //handler.Dispose();
            }

            {
                ///设置单元格样式，设置单元格值等
                handler = new ExcelHelper(excelFilePath, false);
                handler.OpenOrCreate();
                // 获得 Worksheet 对象
                Xls.Worksheet sheet = handler.GetWorksheet("TestSheet");
                //A1-E5
                Xls.Range range = handler.GetRange(sheet, 1, 1, 5, 5);
                handler.SetRangeFormat(range);
                handler.SetCellValue(sheet, 1, 1, " 测试 ");
                handler.SetCellValue(sheet, 2, 1, " 测试 2");
                range.Font.Bold = true;// 加粗
                handler.Save();
                handler.Dispose();
            }
        }
        #endregion
    }
}

