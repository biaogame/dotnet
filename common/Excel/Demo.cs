//将table转换成excel
MemoryStream ContectZongbuTablesms = ExportExcelFactory.exportToExcel(ContectZongbuTables);
string ContectZongbuTablefilePath = ExportExcelFactory.SaveToFile(ContectZongbuTablesms, c.name + "_" + "合同文档期过期_瓶装气业务" + "_" + DateTime.Now.Millisecond, "overtimeareacontect");


string body = GetExcelContent(ContectZongbuTablefilePath, sheetnames);

//将excel中的sheet装换成html表格
   string GetExcelContent(string filePath, List<string> sheetName)
        {
            ExcelHelper excel = new ExcelHelper(filePath);

            StringBuilder strBuild = new StringBuilder();
            foreach (string name in sheetName)
            {
                DataTable dt = excel.ExcelToDataTable(name, true);
                strBuild.Append("<table width=\"100%\" border=\"0\" cellspacing=\"1\" cellpadding=\"4\" bgcolor=\"#cccccc\" style=\"margin-top: 13px;\"  align=\"center\">");
                strBuild.Append("<tr>");
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    strBuild.Append("<td  style=\"background-color:#ffffff;height:25px;line-height:150%;background:#e9faff !important;text-align:center;\">" + dt.Columns[j] + "</td>");
                }
                strBuild.Append("<tr>");
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    strBuild.Append("<tr>");
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        strBuild.Append("<td  style=\"background-color:#ffffff;height:25px;line-height:150%;\">" + dt.Rows[i][dt.Columns[j]] + "</td>");
                    }

                    strBuild.Append("<tr>");
                }
                strBuild.Append("</table>");
            }
            return strBuild.ToString();
        }