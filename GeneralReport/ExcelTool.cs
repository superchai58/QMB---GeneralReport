using System;
using System.Collections.Generic;
using System.Data;
using System.Reflection;

namespace GeneralReport
{
    public class ExcelTool : IDisposable
    {
        private dynamic xlApp = new object();
        private dynamic xlWorkbook = new object();
        private dynamic xlSheet = new object();

        public ExcelTool()
        {
            xlApp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            if (xlApp == null)
            {
                throw new ApplicationException("初始化Excel失败");
            }
        }

        public void CreateDocument()
        {
            xlWorkbook = xlApp.Workbooks.Add(true);
        }

        public void OpenDocument(string fileName)
        {
            xlWorkbook = xlApp.Workbooks.Add(fileName);
        }
        public List<string> GetSheetList()
        {
            List<string> sheetList = new List<string>();
            for (int i = 0; i < xlWorkbook.Worksheets.Count; i++)
            {
                sheetList.Add(xlWorkbook.Worksheets[i + 1].Name);
            }
            return sheetList;
        }

        public bool HasSheet(string sheetName)
        {
            foreach (var sheet in xlWorkbook.Sheets)
            {
                if (sheet.Name == sheetName)
                {
                    return true;
                }
            }
            return false;
        }

        public void InsertSheet(string sheetName)
        {
            int sheetCount = xlWorkbook.Sheets.Count;
            object After = xlWorkbook.Sheets[sheetCount];
            xlSheet = xlWorkbook.Sheets.Add(Missing.Value, After, Missing.Value, Missing.Value);
            xlSheet.Name = sheetName;
        }

        public bool DeleteSheet(string sheetName)
        {
            foreach (var sheet in xlWorkbook.Sheets)
            {
                if (sheet.Name == sheetName)
                {
                    sheet.Delete();
                    return true;
                }
            }
            return false;
        }

        public void Export(Object[,] objData, int startRow, int startColumn, string[] sheetName = null)
        {
            for (int i = 0; i < sheetName.Length; i++)
            {
                if (string.IsNullOrEmpty(sheetName[i]))
                {
                    xlSheet = xlWorkbook.Sheets[1];
                }
                else
                {
                    xlSheet = xlWorkbook.Sheets[sheetName[i]];
                }
                dynamic range = xlSheet.Range(xlSheet.Cells[startRow, startColumn], xlSheet.Cells[startRow + objData.GetLength(0) - 1, startColumn + objData.GetLength(1) - 1]);
                range.Value = objData;
                //xlSheet.Select();//全选，自适应列宽
                //xlSheet.Columns.AutoFit();
            }
        }

        public void Export(DataTable dt, int startRow, int startColumn, bool IsExportColumnName, string[] sheetName = null)
        {
            int rowCount = dt.Rows.Count;
            int columnCount = dt.Columns.Count;
            object[,] objData;

            if (IsExportColumnName)
            {
                objData = new object[rowCount + 1, columnCount + 1];

                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    objData[0, j] = dt.Columns[j].ColumnName;
                }

                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j < columnCount; j++)
                    {
                        objData[i + 1, j] = dt.Rows[i][j];
                    }
                }
            }
            else
            {
                objData = new object[rowCount, columnCount];

                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j < columnCount; j++)
                    {
                        objData[i, j] = dt.Rows[i][j];
                    }
                }
            }

            Export(objData, startRow, startColumn, sheetName);
        }

        public void SetVisible(bool visible)
        {
            xlApp.Visible = visible;
        }

        public void Save(string fileName)
        {
            if (!string.IsNullOrEmpty(fileName))
            {
                xlApp.DisplayAlerts = false;
                xlWorkbook.SaveAs(fileName);
                xlApp.DisplayAlerts = true;
            }
        }

        public void Dispose()
        {
            xlWorkbook.Close();
            xlApp.Quit();

            xlSheet = null;
            xlWorkbook = null;
            xlApp = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
