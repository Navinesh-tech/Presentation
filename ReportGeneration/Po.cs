using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Configuration;
using System.Data.Odbc;
using System.IO;
using System.Net;
using System.Net.Mail;

namespace EmailReport_AMWIN
{
    public static class PO
    {
        private static readonly string PurchaseOrderReportFileName = ConfigurationManager.AppSettings["PurchaseOrderReportFileName"];
        private static readonly string SubVendorPurchaseOrderReportFileName = ConfigurationManager.AppSettings["SubVendorPurchaseOrderReportFileName"];
        private static readonly string SmtpServer = ConfigurationManager.AppSettings["SmtpServer"];
        private static readonly int SmtpPort = Convert.ToInt32(ConfigurationManager.AppSettings["SmtpPort"]);
        private static readonly string SmtpUser = ConfigurationManager.AppSettings["SmtpUser"];
        private static readonly string SmtpPassword = ConfigurationManager.AppSettings["SmtpPassword"];
        private static readonly string FromMail = ConfigurationManager.AppSettings["FromMail"];
        private static readonly string LogDirectory = ConfigurationManager.AppSettings["PurchaseOrderLogDirectory"];
        private static readonly string LogFilePath = Path.Combine(LogDirectory, "PurchaseReport_" + DateTime.Now.ToString("yyyyMMdd") + ".log");

        public static void ProcessPurchaseOrders(OdbcConnection connection)
        {
            InitializeLog();

            try
            {
                string query = @"
SELECT T0.""DocEntry"", T0.""DocNum"", T0.""CardCode"", T1.""CardName"", T1.""U_ToMail"",T1.""U_CCMail"", T1.""U_SubVen""
FROM OPOR T0
INNER JOIN OCRD T1 ON T0.""CardCode"" = T1.""CardCode""
WHERE T1.""CardType"" = 'S'
AND IFNULL(TO_NVARCHAR(T1.""U_ToMail""), '') <> ''
AND IFNULL(T0.""U_EmailSts"", 'No') = 'No'
AND T0.""DocStatus"" = 'O'AND T0.""DocDate"" > '2026-10-12'";


                using (OdbcCommand command = new OdbcCommand(query, connection))
                using (OdbcDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        int docEntry = Convert.ToInt32(reader["DocEntry"]);
                        string docNum = reader["DocNum"].ToString();
                        string cardName = reader["CardName"].ToString();
                        string toMail = reader["U_ToMail"].ToString();
                        string CCMail = reader["U_CCMail"].ToString();
                        string subVendor = reader["U_SubVen"].ToString().Trim().ToUpper();

                        string reportFileName = subVendor == "Y" ? SubVendorPurchaseOrderReportFileName : PurchaseOrderReportFileName;
                        string reportPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, reportFileName);

                        try
                        {
                            byte[] pdfBytes = GeneratePdf(docEntry, reportPath);

                            using (MemoryStream pdfStream = new MemoryStream(pdfBytes))
                            {
                                SendEmail(toMail, CCMail, docNum, cardName, pdfStream);
                            }

                            UpdateEmailStatus(connection, docEntry);
                            Log("Purchase Order " + docNum + " processed successfully. U_SubVen = " + subVendor + ", Report = " + reportFileName);
                        }
                        catch (Exception ex)
                        {
                            Log("Purchase Order " + docNum + " failed: " + ex.Message);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log("Purchase Order process failed: " + ex.Message);
            }
        }

        private static byte[] GeneratePdf(int docEntry, string reportPath)
        {
            if (!File.Exists(reportPath))
            {
                throw new FileNotFoundException("Crystal Report file not found.", reportPath);
            }

            ReportDocument report = new ReportDocument();

            try
            {
                report.Load(reportPath);

                ConnectionInfo connectionInfo = new ConnectionInfo();
                connectionInfo.ServerName = Program.HanaServer;
                connectionInfo.DatabaseName = Program.HanaDatabase;
                connectionInfo.UserID = Program.HanaUser;
                connectionInfo.Password = Program.HanaPassword;

                foreach (Table table in report.Database.Tables)
                {
                    table.LogOnInfo.ConnectionInfo = connectionInfo;
                    table.ApplyLogOnInfo(table.LogOnInfo);
                }

                report.SetParameterValue("DocEntry", docEntry);

                using (Stream stream = report.ExportToStream(ExportFormatType.PortableDocFormat))
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    stream.CopyTo(memoryStream);
                    return memoryStream.ToArray();
                }
            }
            finally
            {
                report.Close();
                report.Dispose();
            }
        }

        private static void SendEmail(string toMail,string CCMail, string docNum, string cardName, MemoryStream pdfStream)
        {
            using (MailMessage mail = new MailMessage())
            using (SmtpClient smtp = new SmtpClient(SmtpServer, SmtpPort))
            {
                mail.From = new MailAddress(FromMail);
                //mail.To.Add(toMail);

                //if (!string.IsNullOrWhiteSpace(CCMail))
                //{
                //    mail.CC.Add(CCMail);
                //}

                foreach (string email in toMail.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))

                {
                    mail.To.Add(email.Trim());
                }

                if (!string.IsNullOrWhiteSpace(CCMail))
                {
                    foreach (string email in CCMail.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))

                    {
                        mail.CC.Add(email.Trim());
                    }
                }

                mail.Subject = "Purchase Order - " + docNum + " - " + cardName;
                mail.Body = "Dear Sir/Madam,\r\n\r\n" +
                     "Please find attached the Purchase Order from Amwin Machining Pvt. Ltd.\r\n" +
                     "We kindly request you to acknowledge receipt of the PO.\r\n\r\n" +
                     "Please note that any order acknowledgement or queries related to the PO " +
                     "should be sent to the email addresses included in the email loop or as specified in the PO.\r\n\r\n" +
                     "We request you not to reply to purchase@amwin.co.in for any communication regarding this PO.\r\n\r\n" +
                     "Thanks & Regards,\r\n\r\n" +
                     "Amwin Machining Pvt Ltd,\r\n" +
                     "Purchase Dept.";
                mail.Attachments.Add(new Attachment(pdfStream, "PO_" + docNum + ".pdf"));

                smtp.EnableSsl = true;
                smtp.Credentials = new NetworkCredential(SmtpUser, SmtpPassword);
                smtp.Send(mail);
            }
        }

        private static void UpdateEmailStatus(OdbcConnection connection, int docEntry)
        {
            string query = "UPDATE OPOR SET \"U_EmailSts\" = 'Yes' WHERE \"DocEntry\" = ?";

            using (OdbcCommand command = new OdbcCommand(query, connection))
            {
                command.Parameters.AddWithValue("@DocEntry", docEntry);
                command.ExecuteNonQuery();
            }
        }

        private static void InitializeLog()
        {
            Directory.CreateDirectory(LogDirectory);
            File.AppendAllText(LogFilePath, Environment.NewLine + "===== Purchase Order Process Started: " + DateTime.Now + " =====" + Environment.NewLine);
        }

        private static void Log(string message)
        {
            Directory.CreateDirectory(LogDirectory);
            File.AppendAllText(LogFilePath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - " + message + Environment.NewLine);
        }
    }
}
