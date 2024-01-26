using System.Data;

namespace GeneralReport
{
    class Report
    {
        private string fileName;

        public string FileName
        {
            get { return fileName; }
        }

        private ReportInfo rptInfo;
        private DataSet reportData;

        public Report(ReportInfo rpt)
        {
            string cmdText = string.Format("select dbo.FormatDate(getdate(),'{0}')", rpt.DateTimeFormat);
            DataSet ds = SQLHelper.ExcuteText(cmdText, rpt.ConnectionString);
            string datetime = ds.Tables[0].Rows[0][0].ToString();

            if (string.IsNullOrEmpty(datetime))
            {
                this.fileName = string.Format("{0}{1}.xlsx", rpt.SavePath, rpt.ReportName);
            }
            else
            {
                this.fileName = string.Format("{0}{1}-{2}.xlsx", rpt.SavePath, rpt.ReportName, datetime);
            }
            this.rptInfo = rpt;
            this.reportData = SQLHelper.ExcuteSP(rpt.SP, rpt.ConnectionString);

            DataTable dt = reportData.Tables[reportData.Tables.Count - 1];

            if(reportData.Tables.Count>rptInfo.Location.Count 
                && dt.Columns[0].ColumnName.Equals("html",System.StringComparison.CurrentCultureIgnoreCase))
            {
                rptInfo.MailBody= reportData.Tables[reportData.Tables.Count - 1].Rows[0]["Html"].ToString();
            }
        }

        public void WriteExcel()
        {
            ExcelTool et = new ExcelTool();
            et.OpenDocument(rptInfo.Template);
            foreach (ReportLocation location in rptInfo.Location)
            {
                et.Export(reportData.Tables[location.ID], location.StartRow, location.StartColumn, location.IsExportColumnName, location.SheetName);
            }
            et.Save(fileName);
            et.Dispose();
        }
    }
}
