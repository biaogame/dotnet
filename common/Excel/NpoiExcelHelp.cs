 
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
 
 // <summary>
    /// NPOI打印Excel
    /// </summary>
    public class NpoiExcelHelp
    {

        /// 转换  ASCII 14 - 31 ->  -   
        /// </summary>  
        /// <param name="tmp"></param>  
        /// <returns></returns>  
        private static string ReplaceLowOrderASCIICharacters(string tmp)
        {
            StringBuilder info = new StringBuilder();
            foreach (char cc in tmp)
            {
                int ss = (int)cc;
                if (((ss >= 0) && (ss <= 8)) || ((ss >= 11) && (ss <= 12)) || ((ss >= 14) && (ss <= 32)))
                    info.AppendFormat(" ", ss);
                else info.Append(cc);
            }
            return info.ToString();
        }

        public static MemoryStream GenerateXlsxBytes(Dictionary<string,string> dts, Dictionary<string, object> dicConfig = null, List<string> funs = null)
        {
            var workBook = new HSSFWorkbook();
            foreach(var item in dts)
            {
                string tableHtml = Encode(item.Value);
                tableHtml = ReplaceLowOrderASCIICharacters(tableHtml);
                string xml = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" + tableHtml;
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(xml);

                XmlNode table = doc.SelectSingleNode("/table");

                int colspan = 1;
                int rowspan = 1;

                int rowNum;
                int columnNum;
                rowNum = 1;
                columnNum = 1;
                var ws = workBook.CreateSheet(item.Key);
                string mapKey = string.Empty;
                string mergKey = string.Empty;

                int rowCount = table.ChildNodes.Count;
                int colCount = FetchColCount(table.ChildNodes);

                InitSheet(ws, rowCount, colCount);

                bool[,] map = new bool[rowCount + 1, colCount + 1];

                foreach (XmlNode row in table.ChildNodes)
                {
                    columnNum = 1;
                    foreach (XmlNode column in row.ChildNodes)
                    {
                        if (column.Attributes["rowspan"] != null)
                        {
                            rowspan = Convert.ToInt32(column.Attributes["rowspan"].Value);
                        }
                        else
                        {
                            rowspan = 1;
                        }

                        if (column.Attributes["colspan"] != null)
                        {
                            colspan = Convert.ToInt32(column.Attributes["colspan"].Value);
                        }
                        else
                        {
                            colspan = 1;
                        }

                        while (map[rowNum, columnNum])
                        {
                            columnNum++;
                        }

                        if (rowspan == 1 && colspan == 1)
                        {
                            SetCellValue(ws, string.Format("{0}{1}", Chr(columnNum), rowNum), column.InnerText);
                            map[rowNum, columnNum] = true;
                        }
                        else
                        {
                            SetCellValue(ws, string.Format("{0}{1}", Chr(columnNum), rowNum), column.InnerText);
                            mergKey =
                                string.Format("{0}{1}:{2}{3}",
                                    Chr(columnNum), rowNum, Chr(columnNum + colspan - 1), rowNum + rowspan - 1);
                            MergCells(ws, mergKey);

                            for (int m = 0; m < rowspan; m++)
                            {
                                for (int n = 0; n < colspan; n++)
                                {
                                    map[rowNum + m, columnNum + n] = true;
                                }
                            }
                        }
                        columnNum++;
                    }
                    rowNum++;
                }

                
                if (dicConfig != null) { MergedRegion(ws, rowCount, colCount, dicConfig); }//根据条件合并指定的单元格
                SheetHeadStyle(ws, workBook, rowCount, colCount);//设置
                                                                 //调用函数
                if (funs != null)
                {
                    for (int i = 0; i < funs.Count; i++)
                    {
                        switch (funs[i])
                        {
                            case "SetBgColor":
                                SetBgColor(ws, workBook, rowCount, colCount); break;
                        }
                    }
                }
                //
                SetDouble(ws, rowCount, colCount);
                SetDate(ws, workBook, rowCount, colCount);
            }
            MemoryStream stream = new MemoryStream();
            workBook.Write(stream);

            return stream;

        }

        /// <summary>
        /// 表头及内容样式
        /// </summary>
        /// <param name="sheet"></param>
        static void SheetHeadStyle(ISheet ws, HSSFWorkbook workBook, int rowCount, int colCount)
        {
            //
            ws.AutoSizeColumn(0);//标题行自适应宽度
            ICellStyle styleHeadNull = workBook.CreateCellStyle();//为了....
            IFont fontHeadNull = workBook.CreateFont();///
            ICellStyle styleHead = workBook.CreateCellStyle();
            IFont fontHead = workBook.CreateFont();//新建一个字体样式对象
            for (int i = 0; i < rowCount; i++)
            {
                IRow rowHead = ws.GetRow(i);//在工作表中：建立行，参数为行号，从0计
                rowHead.Height = Convert.ToInt16(rowHead.Height + (2 * 20));
                for (int j = 0; j < colCount; j++)
                {

                    if (i == 0)
                    {
                        styleHead.Alignment = HorizontalAlignment.Center;//设置单元格的样式：水平对齐居中
                        styleHead.VerticalAlignment = VerticalAlignment.Top;
                        fontHead.Boldweight = short.MaxValue;//设置字体加粗样式
                        fontHead.Color = NPOI.HSSF.Util.HSSFColor.White.Index;
                        //
                        styleHead.BorderBottom = BorderStyle.Thin;//边框
                        styleHead.BorderLeft = BorderStyle.Thin;
                        styleHead.BorderRight = BorderStyle.Thin;
                        styleHead.BorderTop = BorderStyle.Thin;
                        //
                        styleHead.FillPattern = FillPattern.SolidForeground;
                        styleHead.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightBlue.Index;
                    }
                    else
                    {
                        styleHead = styleHeadNull;
                        fontHead = fontHeadNull;//新建一个字体样式对象
                        styleHead.Alignment = HorizontalAlignment.Left;//设置单元格的样式：水平对齐居中
                        styleHead.VerticalAlignment = VerticalAlignment.Top;
                        fontHead.Color = NPOI.HSSF.Util.HSSFColor.Grey80Percent.Index;

                        //
                        styleHead.BorderBottom = BorderStyle.Thin;//边框
                        styleHead.BorderLeft = BorderStyle.Thin;
                        styleHead.BorderRight = BorderStyle.Thin;
                        styleHead.BorderTop = BorderStyle.Thin;
                        //
                    }

                    styleHead.SetFont(fontHead); //使用SetFont方法将字体样式添加到单元格样式中 
                    ICell cellHead = rowHead.GetCell(j);//在行中：建立单元格，参数为列号，从0计
                    cellHead.CellStyle = styleHead;//将新的样式赋给单元格
                    int length = Encoding.Default.GetBytes(cellHead.ToString()).Length;
                    if (length > 250) { length = 200; }
                    if (ws.GetColumnWidth(j) / 256 < length + 1)
                    {
                        ws.SetColumnWidth(j, (length + 2) * 256);
                    }

                }

            }



            //
        }

        /// <summary>
        /// 填充背景色,K3库存导出
        /// </summary>
        static void SetBgColor(ISheet ws, HSSFWorkbook workbook, int rowCount, int colCount)
        {
            ICellStyle styleRed = workbook.CreateCellStyle();
            styleRed.FillPattern = FillPattern.SolidForeground;
            styleRed.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Red.Index;
            styleRed.BorderBottom = BorderStyle.Thin;//边框
            styleRed.BorderLeft = BorderStyle.Thin;
            styleRed.BorderRight = BorderStyle.Thin;
            styleRed.BorderTop = BorderStyle.Thin;
            styleRed.Alignment = HorizontalAlignment.Center;//设置单元格的样式：水平对齐居中
            styleRed.VerticalAlignment = VerticalAlignment.Top;

            ICellStyle styleYellow = workbook.CreateCellStyle();
            styleYellow.FillPattern = FillPattern.SolidForeground;
            styleYellow.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            styleYellow.BorderBottom = BorderStyle.Thin;//边框
            styleYellow.BorderLeft = BorderStyle.Thin;
            styleYellow.BorderRight = BorderStyle.Thin;
            styleYellow.BorderTop = BorderStyle.Thin;
            styleYellow.Alignment = HorizontalAlignment.Center;//设置单元格的样式：水平对齐居中
            styleYellow.VerticalAlignment = VerticalAlignment.Top;
            //
            for (int i = 0; i < rowCount; i++)
            {
                IRow rowNow = ws.GetRow(i);//在工作表中：建立行，参数为行号，从0计
                ICell cellLast = rowNow.GetCell(colCount - 1);
                double result;
                if (double.TryParse(cellLast.StringCellValue.Trim(), out result))
                {
                    if (Convert.ToDouble(cellLast.StringCellValue.Trim()) < 10 && Convert.ToDouble(cellLast.StringCellValue.Trim()) >= 5)
                    {
                        for (int j = 0; j < colCount; j++)
                        {
                            ICell cellNow = rowNow.GetCell(j);
                            cellNow.CellStyle = styleYellow;
                        }
                    }
                    else if (Convert.ToDouble(cellLast.StringCellValue.Trim()) < 5)
                    {
                        for (int j = 0; j < colCount; j++)
                        {
                            ICell cellNow = rowNow.GetCell(j);
                            cellNow.CellStyle = styleRed;
                        }
                    }

                }
            }
        }

        /// <summary>
        /// 根据条件合并指定的单元格
        /// </summary>
        static void MergedRegion(ISheet ws, int rowCount, int colCount, Dictionary<string, object> dicConfig)
        {
            // try{
            int ColIndex = Convert.ToInt32(dicConfig["ColIndex"].ToString());//根据那个列
            int StartIndex = Convert.ToInt32(dicConfig["StartIndex"].ToString());//从哪列开始
            int EndIndex = Convert.ToInt32(dicConfig["EndIndex"].ToString());//从哪列结束
            for (int i = 0; i < rowCount; i++)
            {
                int itemIndex = i;
                IRow rowNow = ws.GetRow(i);//在工作表中：建立行，参数为行号，从0计
                ICell cellNow = rowNow.GetCell(ColIndex);//在行中：建立单元格，参数为列号，从0计
                for (int j = i + 1; j < rowCount; j++)
                {
                    IRow rowNex = ws.GetRow(j);//在工作表中：建立行，参数为行号，从0计
                    ICell cellNex = rowNex.GetCell(ColIndex);//在行中：建立单元格，参数为列号，从0计
                    if (cellNow.StringCellValue == cellNex.StringCellValue)
                    {
                        itemIndex++;
                    }
                    else
                    {
                        if (itemIndex > i)
                        {
                            for (int k = StartIndex; k <= EndIndex; k++)
                            {
                                ws.AddMergedRegion(new CellRangeAddress(i, itemIndex, k, k));
                            }
                            i = itemIndex;
                        }
                        break;
                    }
                    if (j == rowCount - 1)
                    {
                        if (itemIndex > i)
                        {
                            for (int k = StartIndex; k <= EndIndex; k++)
                            {
                                ws.AddMergedRegion(new CellRangeAddress(i, itemIndex, k, k));
                            }
                            i = itemIndex;
                        }
                        break;
                    }
                }

            }

            // }catch { }
        }

        /// <summary>
        /// 把数字转成数值类型
        /// </summary>
        static void SetDouble(ISheet ws, int rowCount, int colCount)
        {
            for (int i = 0; i < rowCount; i++)
            {
                IRow rowNow = ws.GetRow(i);//在工作表中：建立行，参数为行号，从0计
                for (int j = 0; j < colCount; j++)
                {
                    ICell cellNow = rowNow.GetCell(j);//在行中：建立单元格，参数为列号，从0计
                    double result;
                    if (cellNow.StringCellValue.Trim().Length < 15)
                    {
                        if (double.TryParse(cellNow.StringCellValue.Trim(), out result))
                        {
                            cellNow.SetCellValue(Convert.ToDouble(cellNow.StringCellValue.Trim()));
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 把数字转成日期
        /// </summary>
        static void SetDate(ISheet ws, HSSFWorkbook workbook, int rowCount, int colCount)
        {
            //设置单元格格式
            NPOI.HSSF.UserModel.HSSFCellStyle style = (HSSFCellStyle)workbook.CreateCellStyle();
            NPOI.HSSF.UserModel.HSSFDataFormat format = (HSSFDataFormat)workbook.CreateDataFormat();
            style.DataFormat = format.GetFormat("yyyy/mm/dd hh:mm:ss");
            //
            IFont fontHeadNull = workbook.CreateFont();///
            style.Alignment = HorizontalAlignment.Center;//设置单元格的样式：水平对齐居中
            style.VerticalAlignment = VerticalAlignment.Top;
            fontHeadNull.Color = NPOI.HSSF.Util.HSSFColor.Grey80Percent.Index;
            //
            style.BorderBottom = BorderStyle.Thin;//边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //
            for (int i = 0; i < rowCount; i++)
            {
                IRow rowNow = ws.GetRow(i);//在工作表中：建立行，参数为行号，从0计
                for (int j = 0; j < colCount; j++)
                {
                    ICell cellNow = rowNow.GetCell(j);//在行中：建立单元格，参数为列号，从0计
                    DateTime result;
                    if (cellNow.CellType != CellType.Numeric)
                        if (DateTime.TryParse(cellNow.StringCellValue.Trim(), out result))
                        {
                            cellNow.SetCellValue(Convert.ToDateTime(cellNow.StringCellValue.Trim()));

                            cellNow.CellStyle = style;
                        }
                }
            }
        }

        /// <summary>
        /// 过滤HTML标签
        /// </summary>
        /// <param name="str">要过滤的内容</param>
        /// <returns></returns>
        public static string Encode(string str)
        {
            string strTmp = str.Trim();
            //判断字符串不为空
            if (str.Trim().Length > 0)
            {
                //开始过滤字符
                strTmp = Regex.Replace(strTmp, "&", "&amp;");
            }
            return strTmp;
        }

        static int FetchColCount(XmlNodeList nodes)
        {
            int colCount = 0;

            foreach (XmlNode row in nodes)
            {
                if (colCount < row.ChildNodes.Count)
                {
                    colCount = row.ChildNodes.Count;
                }
            }

            return colCount;
        }

        static void InitSheet(ISheet sheet, int rowCount, int colCount)
        {
            for (int i = 0; i < rowCount; i++)
            {
                IRow row = sheet.CreateRow(i);
                for (int j = 0; j < colCount; j++)
                {
                    row.CreateCell(j);
                }
            }
        }

        static void SetCellValue(ISheet sheet, string cellReferenceText, string value)
        {
            CellReference cr = new CellReference(cellReferenceText);
            IRow row = sheet.GetRow(cr.Row);
            ICell cell = row.GetCell(cr.Col);
            cell.SetCellValue(value);

        }

        static void MergCells(ISheet sheet, string mergeKey)
        {
            string[] cellReferences = mergeKey.Split(':');

            CellReference first = new CellReference(cellReferences[0]);
            CellReference last = new CellReference(cellReferences[1]);

            CellRangeAddress region = new CellRangeAddress(first.Row, last.Row, first.Col, last.Col);
            sheet.AddMergedRegion(region);
        }

        public static string Chr(int i)
        {
            if (i > 52)
            {
                char c = (char)(64 + (i - 52));
                return "B" + c.ToString();
            }
            else if (i > 26)
            {
                char c = (char)(64 + (i - 26));
                return "A" + c.ToString();
            }

            else
            {
                char c = (char)(64 + i);
                return c.ToString();
            }

        }
    }