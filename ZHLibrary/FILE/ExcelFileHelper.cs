using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace ZHLibrary.FILE
{
    #region 自定义MemoryStream,用于ExcelFileHelper

    /// <summary>
    /// 为了防止stream被释放，所以单独重写 MemoryStream
    /// 用于导出excel的帮助类 
    /// </summary>
    public class NpoiMemoryStream : MemoryStream
    {
        public NpoiMemoryStream()
        {
            AllowClose = true;
        }

        //手动指定内存是否可以被释放
        public bool AllowClose { get; set; }

        public override void Close()
        {
            //防止系统自动释放
            if (AllowClose)
                base.Close();
        }
    }
    #endregion

    #region 单独设置单元格内容类

    /// <summary>
    /// 额外的excel元素,用来单独设置某个单元格
    /// </summary>
    public class ExtraExcelElemet
    {
        /// <summary>
        /// excel单元格内的文本
        /// </summary>
        public string Text { get; set; }

        /// <summary>
        /// excel单元格所在的行，从上往下计数，从0开始计数
        /// </summary>
        public int Row { get; set; }

        /// <summary>
        /// excel单元格所在的列，从左往右计数，从0开始计数
        /// </summary>
        public int Column { get; set; }

        /// <summary>
        /// 是否需要合并单元格
        /// </summary>
        public bool NeedMerge { get; set; }

        /// <summary>
        /// 合并结束的行号
        /// </summary>
        public int MergeEndRow { get; set; }

        /// <summary>
        /// 合并结束的列号
        /// </summary>
        public int MergeEndColumn { get; set; }

        public ExtraExcelElemet()
        {
            Text = string.Empty;
            Row = 0;
            Column = 0;
            NeedMerge = false;
            MergeEndColumn = 0;
            MergeEndRow = 0;
        }
    }

    #endregion
    public class ExcelFileHelper
    {
        #region 导出excel

        /// <summary>
        /// 向excel中写入几个单独的数据（暂时用于输出异常)
        /// </summary>
        /// <param name="extraExcelElemetList"></param>
        /// <returns></returns>
        public static MemoryStream ExportExcel(List<ExtraExcelElemet> extraExcelElemetList)
        {
            //生成模板excel
            IWorkbook wk = new XSSFWorkbook();

            //读取当前表数据
            ISheet sheet = wk.CreateSheet("sheet");

            //行和列
            IRow row = null;
            ICell cell = null;

            if (extraExcelElemetList != null)
            {
                foreach (var element in extraExcelElemetList) //遍历要写入excel中的元素
                {
                    //创建行和列
                    row = sheet.CreateRow(element.Row);
                    cell = row.CreateCell(element.Column);
                    //设置单元格数值
                    cell.SetCellValue(element.Text);
                }
            }

            //使用自定义的MemoryStream 防止内存被提前释放
            NpoiMemoryStream memoryStream = new NpoiMemoryStream();
            memoryStream.AllowClose = false;

            //将excel数据流保存到stream中
            wk.Write(memoryStream);
            memoryStream.Flush();

            memoryStream.Seek(0, SeekOrigin.Begin);
            memoryStream.AllowClose = true;

            return memoryStream;
        }


        /// <summary>
        /// 通过model list 生成excel数据流
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="modelList">excel的数据源</param>
        /// <param name="startRow">数据起始行，默认为0</param>
        /// <param name="startCol">数据起始列，默认为0</param>
        /// <returns>excel数据流</returns>
        public static MemoryStream ExportExcel<T>(List<T> modelList, int startRow = 0, int startCol = 0)
        {

            //生成模板excel
            IWorkbook wk = new XSSFWorkbook();

            //读取当前表数据
            ISheet sheet = wk.CreateSheet("Sheet");

            //行和列
            IRow row = null;
            ICell cell = null;

            //当前操作的行号
            var currentRow = startRow;

            foreach (var model in modelList) //遍历数据源
            {
                //创建行
                row = sheet.CreateRow(currentRow);
                //获取model的类型
                Type t = typeof(T);
                //当前操作的列号
                var currentColum = startCol;


                //遍历model的参数
                foreach (var propertyInfo in t.GetProperties())
                {
                    //创建列
                    cell = row.CreateCell(currentColum);
                    //获取model的某个字段的value
                    var value = propertyInfo.GetValue(model);
                    //设置单元格的数值
                    cell.SetCellValue(value == null ? "" : value.ToString());
                    currentColum++;
                }
                currentRow++;
            }

            //使用自定义的MemoryStream 防止内存被提前释放
            NpoiMemoryStream memoryStream = new NpoiMemoryStream();
            memoryStream.AllowClose = false;
            wk.Write(memoryStream);
            memoryStream.Flush();

            memoryStream.Seek(0, SeekOrigin.Begin);
            memoryStream.AllowClose = true;

            return memoryStream;
        }

        /// <summary>
        /// 导出excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="modelList"></param>
        /// <param name="extraElemetList"></param>
        /// <param name="startRow"></param>
        /// <param name="startCol"></param>
        /// <param name="useDefaultStyle"></param>
        /// <param name="exportPropertiesList"></param>
        /// <returns></returns>
        public static MemoryStream ExportExcelWithExtra<T>(List<T> modelList,
            List<ExtraExcelElemet> extraElemetList, int startRow = 0, int startCol = 0, bool useDefaultStyle = true, List<string> exportPropertiesList = null)
        {
            //记录列数
            int parmCount = 0;

            //生成模板excel
            IWorkbook wk = new XSSFWorkbook();

            //读取当前表数据
            ISheet sheet = wk.CreateSheet("Sheet1");

            IRow row = null; //读取当前行数据
            ICell cell = null;

            //当前操作的行号
            var currentRow = startRow;

            foreach (var model in modelList) //遍历数据源
            {
                //创建行
                row = sheet.GetRow(currentRow) ?? sheet.CreateRow(currentRow);
                //获取model的类型
                Type t = typeof(T);

                //当前操作的列号
                var currentColum = startCol;

                if (exportPropertiesList == null)
                {
                    parmCount = t.GetProperties().Length;
                }
                else
                {
                    parmCount = exportPropertiesList.Count;
                }

                //遍历model的参数
                foreach (var propertyInfo in t.GetProperties())
                {
                    if (exportPropertiesList == null)
                    {
                        //创建列
                        cell = row.GetCell(currentColum) ?? row.CreateCell(currentColum);
                        //获取model的某个字段的value
                        var value = propertyInfo.GetValue(model);
                        //设置单元格的数值
                        cell.SetCellValue(value == null ? "" : value.ToString());
                        if (useDefaultStyle)
                        {
                            cell.CellStyle = GetDetaulCellStyle(wk);
                        }

                        currentColum++;
                    }
                    else
                    {
                        if (exportPropertiesList.Contains(propertyInfo.Name))//获取字段名称
                        {
                            //创建列
                            cell = row.GetCell(currentColum) ?? row.CreateCell(currentColum);
                            //获取model的某个字段的value
                            var value = propertyInfo.GetValue(model);
                            //设置单元格的数值
                            cell.SetCellValue(value == null ? "" : value.ToString());
                            if (useDefaultStyle)
                            {
                                cell.CellStyle = GetDetaulCellStyle(wk);
                            }

                            currentColum++;
                        }
                        else if (propertyInfo.Name == "JustOneElement")
                        {
                            //获取model的某个字段的value
                            var value = propertyInfo.GetValue(model);
                            //合并单元格
                            if ((bool)value)
                            {
                                var region = new CellRangeAddress(currentRow, currentRow, 0, parmCount - 1);
                                sheet.AddMergedRegion(region);
                                SetRegionAllBorder(region, sheet, wk);
                            }

                        }
                    }
                }
                currentRow++;
            }


            if (extraElemetList != null)
            {
                foreach (var element in extraElemetList) //遍历额外添加的元素
                {
                    //读取或创建行
                    row = sheet.GetRow(element.Row) ?? sheet.CreateRow(element.Row);
                    //读取或创建列
                    cell = row.GetCell(element.Column) ?? row.CreateCell(element.Column);
                    //设置单元格的值
                    cell.SetCellValue(element.Text);

                    //判断是否添加样式
                    if (useDefaultStyle)
                    {
                        cell.CellStyle = GetDetaulCellStyle(wk);
                    }

                    //是否需要合并单元格
                    if (element.NeedMerge)
                    {
                        //CellRangeAddress四个参数为：起始行，结束行，起始列，结束列
                        var region = new CellRangeAddress(element.Row, element.MergeEndRow, element.Column, element.MergeEndColumn);
                        sheet.AddMergedRegion(region);
                        SetRegionAllBorder(region, sheet, wk);
                    }
                }
            }

            //自使用列宽
            for (int i = 0; i < parmCount; i++)
            {
                sheet.AutoSizeColumn(i);
            }


            //使用自定义的MemoryStream 防止内存被提前释放
            NpoiMemoryStream memoryStream = new NpoiMemoryStream();
            memoryStream.AllowClose = false;
            wk.Write(memoryStream);
            memoryStream.Flush();

            memoryStream.Seek(0, SeekOrigin.Begin);
            memoryStream.AllowClose = true;

            return memoryStream;
        }

        /// <summary>
        /// 通过指定excel模板和model List生成excel数据流
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="modelList">excel数据源</param>
        /// <param name="templetExcelPath">excel模板名称，在config文件中配置</param>
        /// <param name="startRow">数据起始行，默认为0</param>
        /// <param name="startCol">数据起始列，默认为0</param>
        /// <param name="exportPropertiesList">需要导出的参数List，默认为空则会导出model的所有字段，如果只需要导出model的指定字段，则在此list中添加字段的名称</param>
        /// <returns></returns>
        public static MemoryStream ExportExcelFromTempler<T>(List<T> modelList, string templetExcelPath, int startRow = 0,
            int startCol = 0, List<string> exportPropertiesList = null)
        {
            //判断文件存在
            if (!File.Exists(templetExcelPath))
            {
                throw new Exception(string.Format("无法找到模板:{0}", templetExcelPath));
            }

            FileStream fileStream = new FileStream(templetExcelPath, FileMode.Open);
            //从内存中获取模板
            byte[] templetByte = new byte[fileStream.Length];
            fileStream.Read(templetByte, 0, templetByte.Length);
            // 设置当前流的位置为流的开始
            fileStream.Seek(0, SeekOrigin.Begin);

            //生成模板excel
            IWorkbook wk = new XSSFWorkbook(new MemoryStream(templetByte));

            //读取当前表数据
            ISheet sheet = wk.GetSheetAt(0);

            IRow row = null;
            ICell cell = null;

            //当前操作的行号
            var currentRow = startRow;


            foreach (var model in modelList) //遍历数据源
            {
                //如果模板中存在当前行，则读取当前行，否则创建一个新行
                row = sheet.GetRow(currentRow) ?? sheet.CreateRow(currentRow);
                //获取model的类型
                Type t = typeof(T);
                //当前操作的列号
                var currentColum = startCol;

                if (exportPropertiesList == null) //判断是否需要填充指定字段
                {
                    //填充model的所有字段
                    //遍历model的参数
                    foreach (var propertyInfo in t.GetProperties())
                    {
                        //读取或创建列
                        cell = row.GetCell(currentColum) ?? row.CreateCell(currentColum);
                        //获取model的某个字段的value
                        var value = propertyInfo.GetValue(model);
                        //设置单元格的数值
                        cell.SetCellValue(value == null ? "" : value.ToString());
                        currentColum++;
                    }
                }
                else
                {
                    foreach (var propertyInfo in t.GetProperties())
                    {
                        if (exportPropertiesList.Contains(propertyInfo.Name)) //填充model的指定字段
                        {
                            //读取或创建列
                            cell = row.GetCell(currentColum) ?? row.CreateCell(currentColum);
                            //获取model的某个字段的value
                            var value = propertyInfo.GetValue(model);
                            //设置单元格的数值
                            cell.SetCellValue(value == null ? "" : value.ToString());
                            currentColum++;
                        }
                    }
                }
                currentRow++;
            }

            //使用自定义的MemoryStream 防止内存被提前释放
            NpoiMemoryStream memoryStream = new NpoiMemoryStream();
            memoryStream.AllowClose = false;
            wk.Write(memoryStream);
            memoryStream.Flush();

            memoryStream.Seek(0, SeekOrigin.Begin);
            memoryStream.AllowClose = true;

            return memoryStream;
        }


        /// <summary>
        /// 根据model List和指定导出字段来生成excel数据流
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="modelList">数据集</param>
        /// <param name="exportPropertiesList">需要导出的参数List，默认为空则会导出model的所有字段，如果只需要导出model的指定字段，则在此list中添加字段的名称</param>
        /// <param name="startRow">数据起始行，默认为0</param>
        /// <param name="startCol">数据起始列，默认为0</param>>
        /// <returns></returns>
        public static MemoryStream ExportExcelWithPropertyList<T>(List<T> modelList, List<string> exportPropertiesList,
            int startRow = 0, int startCol = 0)
        {

            //生成模板excel
            IWorkbook wk = new XSSFWorkbook();

            //读取当前表数据
            ISheet sheet = wk.CreateSheet("Sheet");

            IRow row = null;
            ICell cell = null;

            //当前操作的行号
            var currentRow = startRow;
            //当前操作的列号
            var currentColum = startCol;

            //遍历model list
            foreach (var model in modelList)
            {
                //创建行
                row = sheet.CreateRow(currentRow);
                //获取model类型
                Type t = typeof(T);

                foreach (var propertyInfo in t.GetProperties()) //遍历model的字段
                {
                    if (exportPropertiesList.Contains(propertyInfo.Name)) //填充model的指定字段
                    {
                        //创建列
                        cell = row.GetCell(currentColum) ?? row.CreateCell(currentColum);
                        //获取model的某个字段的value
                        var value = propertyInfo.GetValue(model);
                        //设置单元格的数值
                        cell.SetCellValue(value == null ? "" : value.ToString());
                        currentColum++;
                    }
                }
                currentRow++;
            }

            //使用自定义的MemoryStream 防止内存被提前释放
            NpoiMemoryStream memoryStream = new NpoiMemoryStream();
            memoryStream.AllowClose = false;
            wk.Write(memoryStream);
            memoryStream.Flush();

            memoryStream.Seek(0, SeekOrigin.Begin);
            memoryStream.AllowClose = true;

            return memoryStream;
        }


        /// <summary>
        /// 通过指定模板和单独添加的单元格和model List来生成excel数据流
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="modelList">数据集</param>
        /// <param name="extraElemetList">需要单独添加的数据，如需要在0，0位置添加一个单独的标题</param>
        /// <param name="templetName">excel模板名称，在config文件中配置</param>
        /// <param name="startRow">数据起始行，默认为0</param>
        /// <param name="startCol">数据起始列，默认为0</param>>
        /// <returns></returns>
        public static MemoryStream ExportExcelFromTempletWithExtra<T>(List<T> modelList,
            List<ExtraExcelElemet> extraElemetList, string templetName, int startRow = 0, int startCol = 0)
        {
            MemoryStream tempStream = ExportExcelFromTempler(modelList, templetName, startRow, startCol);
            //生成模板excel
            IWorkbook wk = new XSSFWorkbook(tempStream);

            //读取当前表数据
            ISheet sheet = wk.GetSheetAt(0);

            IRow row = null; //读取当前行数据
            ICell cell = null;

            //当前操作的行号
            var currentRow = startRow;

            foreach (var model in modelList) //遍历数据源
            {
                //创建行
                row = sheet.GetRow(currentRow) ?? sheet.CreateRow(currentRow);
                //获取model的类型
                Type t = typeof(T);

                //当前操作的列号
                var currentColum = startCol;

                //遍历model的参数
                foreach (var propertyInfo in t.GetProperties())
                {
                    //创建列
                    cell = row.GetCell(currentColum) ?? row.CreateCell(currentColum);
                    //获取model的某个字段的value
                    var value = propertyInfo.GetValue(model);
                    //设置单元格的数值
                    cell.SetCellValue(value == null ? "" : value.ToString());
                    currentColum++;
                }
                currentRow++;
            }


            if (extraElemetList != null)
            {
                foreach (var element in extraElemetList) //遍历额外添加的元素
                {
                    //读取或创建行
                    row = sheet.GetRow(element.Row) ?? sheet.CreateRow(element.Row);
                    //读取或创建列
                    cell = row.GetCell(element.Column) ?? row.CreateCell(element.Column);
                    //设置单元格的值
                    cell.SetCellValue(element.Text);
                }
            }

            //使用自定义的MemoryStream 防止内存被提前释放
            NpoiMemoryStream memoryStream = new NpoiMemoryStream();
            memoryStream.AllowClose = false;
            wk.Write(memoryStream);
            memoryStream.Flush();

            memoryStream.Seek(0, SeekOrigin.Begin);
            memoryStream.AllowClose = true;

            return memoryStream;
        }

        #endregion

        #region excel样式
        /// <summary>
        /// 获取默认单元格样式
        /// </summary>
        /// <param name="wk"></param>
        /// <returns></returns>
        private static ICellStyle GetDetaulCellStyle(IWorkbook wk)
        {
            ICellStyle cellStyle = wk.CreateCellStyle();

            //设置单元格上下左右边框线
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;

            //文字水平和垂直对齐方式  
            cellStyle.Alignment = HorizontalAlignment.Center;
            cellStyle.VerticalAlignment = VerticalAlignment.Center;

            //自动换行
            cellStyle.WrapText = true;
            return cellStyle;
        }

        /// <summary>
        /// 设置合并区域边框
        /// </summary>
        /// <param name="region">合并区域</param>
        /// <param name="sheet">sheet</param>
        /// <param name="wk">excel文件</param>
        private static void SetRegionAllBorder(CellRangeAddress region, ISheet sheet, IWorkbook wk)
        {
            RegionUtil.SetBorderTop(1, region, sheet, wk);
            RegionUtil.SetBorderBottom(1, region, sheet, wk);
            RegionUtil.SetBorderLeft(1, region, sheet, wk);
            RegionUtil.SetBorderRight(1, region, sheet, wk);
        }

        #endregion
    }
}
