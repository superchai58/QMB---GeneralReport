using System.Collections.Generic;
using System.Xml;

namespace GeneralReport
{
    class XMLTool
    {
        public static List<ReportLocation> FormatXMLData(string xmlData)
        {
            List<ReportLocation> list = new List<ReportLocation>();

            if (string.IsNullOrEmpty(xmlData) == false)
            {
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(xmlData);

                XmlNode xn = doc.SelectSingleNode("Location");

                XmlNodeList xnl = xn.ChildNodes;

                foreach (XmlNode item in xnl)
                {
                    XmlElement xe = (XmlElement)item;
                    list.Add(new ReportLocation()
                    {
                        ID = int.Parse(xe.GetAttribute("ID")),
                        SheetName = xe.GetAttribute("SheetName").Split(';'),
                        StartRow = int.Parse(xe.GetAttribute("StartRow")),
                        StartColumn = int.Parse(xe.GetAttribute("StartColumn")),
                        IsExportColumnName = xe.GetAttribute("IsExportColumnName").ToLower() == "true" ? true : false
                    }
                    );
                }
            }
            return list;
        }
    }
}
