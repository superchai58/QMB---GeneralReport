using System.Collections.Generic;

namespace GeneralReport
{
    class ReportInfo
    {
        private string connectionString;

        public string ConnectionString
        {
            get { return connectionString; }
            set { connectionString = value; }
        }
        
        private string reportName;

        public string ReportName
        {
            get { return reportName; }
            set { reportName = value; }
        }

        private List<ReportLocation> location;

        public List<ReportLocation> Location
        {
            get { return location; }
            set { location = value; }
        }

        private string sp;

        public string SP
        {
            get { return sp; }
            set { sp = value; }
        }

        private string template;

        public string Template
        {
            get { return template; }
            set { template = value; }
        }

        private string savePath;

        public string SavePath
        {
            get { return savePath; }
            set { savePath = value; }
        }

        private string dateTimeFormat;

        public string DateTimeFormat
        {
            get { return dateTimeFormat; }
            set { dateTimeFormat = value; }
        }

        public string MailBody { get; set; }
    }
}
