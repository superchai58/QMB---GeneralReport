using Connect.BLL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace GeneralReport
{
    class Program
    {
        static void Main(string[] args)
        {
            //string connectionString = "Data Source=172.26.40.20;Initial Catalog=Report;UID=sa;PWD=pqms#9vd";
            string connectionString = "Data Source=10.97.1.12;Initial Catalog=Report;UID=sa;PWD=pqmb#7sa";      //superchai modify 20230512
            string timeGroup = "";
            string reportName = "";
            //string timeGroup = "0758";
            //string reportName = "QIMS_InputOutptDetailSN";
            

            try
            {
                //if (args.Length > 0)      //superchai modify 20230512
                //{     //superchai modify 20230512
                    //----------------superchai modify 20230512 GettimeGroup & reportName (Begin)----------------------
                    DataTable dt = new DataTable();
                    SqlCommand cmd = new SqlCommand();
                    ConnectDB oCon = new ConnectDB();

                    cmd.CommandText = "SELECT [ReportName],[TimeGroup] FROM [ReportDefine] with(nolock) Where TimeGroup = '0815' Order by TimeGroup";
                    cmd.CommandTimeout = 180;
                    dt = oCon.Query(cmd);
                    if (dt.Rows.Count > 0)
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            timeGroup = row["TimeGroup"].ToString().Trim();
                            reportName = row["ReportName"].ToString().Trim();

                            //----------------superchai modify 20230512 GettimeGroup & reportName (End)----------------------

                            //--------superchai modify 20230512 (Begin)--------------
                            //timeGroup = args[0];

                            //if (args.Length > 1)//手动执行指定报表
                            //{
                            //    reportName = args[1];
                            //}
                            //--------superchai modify 20230512 (Begin)--------------
                            List<ReportInfo> list = GetAllReportInfo(connectionString, timeGroup, reportName);
                            foreach (ReportInfo rptInfo in list)
                            {
                                try
                                {
                                    Report rpt = new Report(rptInfo);
                                    //Console.WriteLine(rptInfo.ReportName);
                                    rpt.WriteExcel();
                                    SendMail(connectionString: connectionString, reportName: rptInfo.ReportName, attachment: rpt.FileName, mailBody: rptInfo.MailBody);
                                }
                                catch (Exception ex)
                                {
                                    SendMail(connectionString: connectionString, reportName: rptInfo.ReportName, mailBody: ex.ToString());
                                }
                            }
                        }                        
                    }
                    
                //}
            }
            catch (Exception ex1)
            {
                SendExceptionMail(connectionString: connectionString, errMsg: ex1.ToString());
            }
        }

        public static List<ReportInfo> GetAllReportInfo(string connectionString, string timeGroup, string reportName)
        {
            string cmdText;
            if (string.IsNullOrEmpty(reportName))
            {
                cmdText = string.Format("SELECT B.A,* FROM ReportDefine A CROSS APPLY DBO.func_split(A.TimeGroup,';') B WHERE B.a='{0}'", timeGroup);
            }
            else
            {
                //手动执行指定报表
                cmdText = string.Format("select * from ReportDefine where ReportName='{0}'", reportName);
            }
            DataSet ds = SQLHelper.ExcuteText(cmdText, connectionString);
            DataTable dt = ds.Tables[0];
            List<ReportInfo> list = new List<ReportInfo>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                list.Add(new ReportInfo()
                {
                    ReportName = dt.Rows[i]["ReportName"].ToString(),
                    ConnectionString = dt.Rows[i]["ConnectionString"].ToString(),
                    Location = XMLTool.FormatXMLData(dt.Rows[i]["Location"].ToString()),
                    SP = dt.Rows[i]["SP"].ToString(),
                    Template = dt.Rows[i]["Template"].ToString(),
                    SavePath = dt.Rows[i]["SavePath"].ToString(),
                    DateTimeFormat = dt.Rows[i]["DateTimeFormat"].ToString(),
                    MailBody=dt.Rows[i]["MailSubject"].ToString()
                }
                );
            }
            return list;
        }

        public static void SendMail(string connectionString, string reportName, string attachment = "", string mailBody = "")
        {
            string cmdText = string.Format("exec GeneralReport_SendMail @ReportName='{0}',@Attachment='{1}',@ErrMsg=N'{2}'", reportName, attachment, mailBody);
            SQLHelper.ExcuteText(cmdText, connectionString);
        }

        public static void SendExceptionMail(string connectionString, string errMsg = "")
        {
            string cmdText = string.Format("exec GeneralReport_SendExceptionMail @ErrMsg=N'{0}'", errMsg);
            SQLHelper.ExcuteText(cmdText, connectionString);
        }
    }
}
