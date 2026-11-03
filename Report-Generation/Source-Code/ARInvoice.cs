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
    public static class AR_Invoice
    {
        private static readonly string InternalReportFileName = ConfigurationManager.AppSettings["InternalReportFileName"];
        private static readonly string ExternalReportFileName = ConfigurationManager.AppSettings["ExternalReportFileName"];
        private static readonly string SmtpServer = ConfigurationManager.AppSettings["SmtpServer"];
        private static readonly int SmtpPort = Convert.ToInt32(ConfigurationManager.AppSettings["SmtpPort"]);
        private static readonly string SmtpUser = ConfigurationManager.AppSettings["SmtpUser"];
        private static readonly string SmtpPassword = ConfigurationManager.AppSettings["SmtpPassword"];
        private static readonly string FromMail = ConfigurationManager.AppSettings["FromMail"];
        private static readonly string LogDirectory = ConfigurationManager.AppSettings["ARInvoiceLogDirectory"];
        private static readonly string LogFilePath = Path.Combine(LogDirectory, "ARInvoiceReport_" + DateTime.Now.ToString("yyyyMMdd") + ".log");

        public static void ProcessARInvoices(OdbcConnection connection)
        {
            InitializeLog();

            try
            {
                string query = @"
SELECT T0.""DocEntry"", T0.""DocNum"",T0.""DocDate"",T0.""DocDueDate"",T0.""DocTotal"", T0.""CardCode"", T1.""CardName"", T1.""U_ToMail"",T1.""U_CCMail"", T0.""U_IRN"", T0.""U_OrdrType""
FROM OINV T0
INNER JOIN OCRD T1 ON T0.""CardCode"" = T1.""CardCode""
WHERE T1.""CardType"" = 'C'
AND IFNULL(T0.""U_IRN"", '') <> ''
AND T0.""U_OrdrType"" IN ('LB', 'LJ', 'EB')
AND IFNULL(TO_NVARCHAR(T1.""U_ToMail""), '') <> ''
AND IFNULL(T0.""U_EmailSts"", 'No') = 'No'
AND T0.""DocStatus"" = 'O' AND T0.""DocDate"" > '2026-10-12'";


                using (OdbcCommand command = new OdbcCommand(query, connection))
                using (OdbcDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        int docEntry = Convert.ToInt32(reader["DocEntry"]);
                        string docNum = reader["DocNum"].ToString();
                        string cardName = reader["CardName"].ToString();
                        string invoiceDate = Convert.ToDateTime(reader["DocDate"]).ToString("dd-MM-yyyy");

                        string totalAmount = Convert.ToDecimal(reader["DocTotal"]).ToString("N2");

                        string dueDate = Convert.ToDateTime(reader["DocDueDate"]).ToString("dd-MM-yyyy");
                        string toMail = reader["U_ToMail"].ToString();
                        string ccMail = reader["U_CCMail"].ToString();
                        string orderType = reader["U_OrdrType"].ToString().Trim().ToUpper();

                        string reportFileName = orderType == "EB" ? ExternalReportFileName : InternalReportFileName;
                        string reportPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, reportFileName);

                        try
                        {
                            byte[] pdfBytes = GeneratePdf(docEntry, reportPath);

                            using (MemoryStream pdfStream = new MemoryStream(pdfBytes))
                            {
                                SendEmail(toMail, ccMail, docNum, cardName,invoiceDate,totalAmount,dueDate, pdfStream);
                            }

                            UpdateEmailStatus(connection, docEntry);
                            Log("AR Invoice " + docNum + " processed successfully. Order Type = " + orderType + ", Report = " + reportFileName);
                        }
                        catch (Exception ex)
                        {
                            Log("AR Invoice " + docNum + " failed: " + ex.Message);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log("AR Invoice process failed: " + ex.Message);
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

        private static void SendEmail(string toMail, string ccMail, string docNum, string cardName,string invoiceDate,string totalAmount,string dueDate, MemoryStream pdfStream)
        {
            using (MailMessage mail = new MailMessage())
            using (SmtpClient smtp = new SmtpClient(SmtpServer, SmtpPort))
            {
                mail.From = new MailAddress(FromMail);
                
                foreach (string email in toMail.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))

                {
                    mail.To.Add(email.Trim());
                }

                if (!string.IsNullOrWhiteSpace(ccMail))
                {
                    foreach (string email in ccMail.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))

                    {
                        mail.CC.Add(email.Trim());
                    }
                }
                mail.Subject = "AR Invoice - " + docNum + " - " + cardName;

                mail.IsBodyHtml = true;
                mail.Body = "<html><body style='font-family:Arial; font-size:12px;'>" +
                "<p>Dear Sir / Madam,</p>" +
                "<p>Please find attached the invoice for your reference and records.</p>" +
                "<p><b>Invoice Details:</b><br/>" +
                "Invoice Number: " + docNum + "<br/>" +
                "Invoice Date: " + invoiceDate + "<br/>" +
                "Total Amount: " + totalAmount + "<br/>" +
                "Payment Due Date: " + dueDate + "</p>" +
                "<p>Kindly review the attached invoice and arrange payment by the due date.</p>" +
                "<p>Thank you for your business and continued support.</p>" +
                "<br/>" +
                "<p>Thanks and Regards<br/>" +
                //"Finance Department.<br/><br/>" +
                "Amwin Machining Pvt Ltd.,<br/>" +
                //"Ground Floor, Sy.No. 41/2 & 41/3,<br/>" +
                //"Karivobanahalli, Peenya 2nd Stage.<br/>" +
                //"GSTIN :- 29AARCA1318B1ZW</p>" +
                "</body></html>";
                mail.Attachments.Add(new Attachment(pdfStream, "ARInvoice_" + docNum + ".pdf"));

                smtp.EnableSsl = true;
                smtp.Credentials = new NetworkCredential(SmtpUser, SmtpPassword);
                smtp.Send(mail);
            }
        }

        private static void UpdateEmailStatus(OdbcConnection connection, int docEntry)
        {
            string query = "UPDATE OINV SET \"U_EmailSts\" = 'Yes' WHERE \"DocEntry\" = ?";

            using (OdbcCommand command = new OdbcCommand(query, connection))
            {
                command.Parameters.AddWithValue("@DocEntry", docEntry);
                command.ExecuteNonQuery();
            }
        }

        private static void InitializeLog()
        {
            Directory.CreateDirectory(LogDirectory);
            File.AppendAllText(LogFilePath, Environment.NewLine + "===== AR Invoice Process Started: " + DateTime.Now + " =====" + Environment.NewLine);
        }

        private static void Log(string message)
        {
            Directory.CreateDirectory(LogDirectory);
            File.AppendAllText(LogFilePath, DateTime.Now.ToString("yyyy-MM-dd HH:mm") + " - " + message + Environment.NewLine);
        }
    }
}
