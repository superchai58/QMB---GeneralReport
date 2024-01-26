namespace GeneralReport
{
    class ReportLocation
    {
        private int id;

        public int ID
        {
            get { return id; }
            set { id = value; }
        }

        private string[] sheetName;

        public string[] SheetName
        {
            get { return sheetName; }
            set { sheetName = value; }
        }

        private int startRow;

        public int StartRow
        {
            get { return startRow; }
            set { startRow = value; }
        }

        private int startColumn;

        public int StartColumn
        {
            get { return startColumn; }
            set { startColumn = value; }
        }

        private bool isExportColumnName;

        public bool IsExportColumnName
        {
            get { return isExportColumnName; }
            set { isExportColumnName = value; }
        }
    }
}
