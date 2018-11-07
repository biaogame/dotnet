public class ExportExcelFactory
    {
        public static MemoryStream exportToExcel(Dictionary<string,DataTable> dts, Dictionary<string, object> dicConfig = null, List<string> funs = null)
        {
            MemoryStream bytes = null;
            Dictionary<string, string> dtstring = new Dictionary<string, string>();
            foreach(var item in dts)
            {
                dtstring[item.Key] = GetHtml(item.Value);
            }
            bytes = NpoiExcelHelp.GenerateXlsxBytes(dtstring, dicConfig, funs);
            return bytes;
        }

        public static string GetHtml(DataTable dt)
        {
            StringBuilder strHtml = new StringBuilder();
            strHtml.Append("<table>");

            strHtml.Append("<tr>");//表头
            foreach (DataColumn cl in dt.Columns)
            {
                strHtml.Append("<td>" + cl.ColumnName + "</td>");
            }
            strHtml.Append("</tr>");

            for (int i = 0; i < dt.Rows.Count; i++)//表内容
            {
                strHtml.Append("<tr>");
                foreach (DataColumn cl in dt.Columns)
                {
                    strHtml.Append("<td>" + dt.Rows[i][cl.ColumnName].ToString() + "</td>");
                }
                strHtml.Append("</tr>");

            }
            strHtml.Append("</table>");
            return strHtml.ToString();
        }

        public static string SaveToFile(MemoryStream ms, string fileName,string floder)
        {
            if (Directory.Exists("File/" + floder) == false)//如果不存在就创建file文件夹
            {
                Directory.CreateDirectory("File/" + floder);
            }
            string filePath = "File/"+ floder + "/" + fileName + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            using (FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                byte[] data = ms.ToArray();

                fs.Write(data, 0, data.Length);
                fs.Flush();
                data = null;
            }
            return filePath;
        }
    }