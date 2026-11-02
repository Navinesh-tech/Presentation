using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.Odbc;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace EmailReport_AMWIN
{
    public static class GRPO
    {
        private static readonly string SmtpServer = ConfigurationManager.AppSettings["SmtpServer"];
        private static readonly int SmtpPort = Convert.ToInt32(ConfigurationManager.AppSettings["SmtpPort"]);
        private static readonly string SmtpUser = ConfigurationManager.AppSettings["SmtpUser"];
        private static readonly string SmtpPassword = ConfigurationManager.AppSettings["SmtpPassword"];
        private static readonly string FromMail = ConfigurationManager.AppSettings["FromMail"];
        private static readonly string LogDirectory = ConfigurationManager.AppSettings["GRPOLogDirectory"];
        private static readonly string LogFilePath = Path.Combine(LogDirectory, "GRPOReport_" + DateTime.Now.ToString("yyyyMMdd") + ".log");

        public static void ProcessGRPO(OdbcConnection connection)
        {
            InitializeLog();

            try
            {
                string query = @"
SELECT T0.""DocEntry"",
       T0.""DocNum"" as ""GRN No"",
       T0.""DocDate"" as ""GRN Date"",
      T0.""TaxDate"" AS ""Invoice Date"",
       T0.""CardCode"" as ""Vendor Code"",
       T0.""CardName"" as ""Vendor Name"",
       T0.""NumAtCard"" AS ""Invoice No"",
       T1.""ItemCode"",
       T1.""Dscription"",
       T1.""Quantity"",
       T1.""Price"",
       T1.""LineTotal"",
       T2.""U_ToMail"",
       T2.""U_CCMail""
FROM OPDN T0
INNER JOIN PDN1 T1 ON T0.""DocEntry"" = T1.""DocEntry""
INNER JOIN OCRD T2 ON T0.""CardCode"" = T2.""CardCode""
WHERE T0.""CANCELED"" = 'N'
AND IFNULL(T0.""U_EmailSts"", 'No') = 'No'
AND IFNULL(TO_NVARCHAR(T2.""U_ToMail""), '') <> ''
AND T0.""DocStatus"" = 'O'
AND T1.""ItemCode"" <> 'SubconFee'

AND T0.""DocDate"" > '2026-10-12'
ORDER BY T1.""LineNum""";

                using (OdbcCommand command = new OdbcCommand(query, connection))
                using (OdbcDataReader reader = command.ExecuteReader())
                {
                    Dictionary<string, VendorGRPOData> vendors = new Dictionary<string, VendorGRPOData>();
              

                 
                    while (reader.Read())
                    {
                        string cardCode = reader["Vendor Code"].ToString();

                        if (!vendors.ContainsKey(cardCode))

                            if (!vendors.ContainsKey(cardCode))
                            {
                                if (cardCode == "0")
                                {
                                    Log("CardCode is 0.");
                                }

                                VendorGRPOData vendor = new VendorGRPOData();

                                vendor.CardCode = cardCode;
                                vendor.CardName = reader["Vendor Name"].ToString();
                                vendor.ToMail = reader["U_ToMail"].ToString();
                                vendor.CcMail = reader["U_CCMail"].ToString();
                                vendor.HtmlData = new StringBuilder();

                                vendor.HtmlData.AppendLine(
                                    "<table border='1' cellpadding='5' cellspacing='0' " +
                                    "style='border-collapse:collapse; width:100%; " +
                                    "font-family:Arial; font-size:12px;'>");

                                vendor.HtmlData.AppendLine(
                                    "<tr style='font-weight:bold; background-color:#D9EAF7;'>");

                                vendor.HtmlData.AppendLine("<th>GRN No</th>");
                                vendor.HtmlData.AppendLine("<th>GRN Date</th>");
                                vendor.HtmlData.AppendLine("<th>Invoice No</th>");
                                vendor.HtmlData.AppendLine("<th>Invoice Date</th>");

                                vendor.HtmlData.AppendLine("<th>Vendor Code</th>");
                                vendor.HtmlData.AppendLine("<th>Vendor Name</th>");
                                vendor.HtmlData.AppendLine("<th>ItemCode</th>");

                                vendor.HtmlData.AppendLine("<th>Description</th>");
                                vendor.HtmlData.AppendLine("<th>Quantity</th>");
                                vendor.HtmlData.AppendLine("<th>Price</th>");
                                vendor.HtmlData.AppendLine("<th>LineTotal</th>");

                                vendor.HtmlData.AppendLine("</tr>");

                                vendors.Add(cardCode, vendor);
                            }

                        VendorGRPOData currentVendor = vendors[cardCode];


                        currentVendor.HtmlData.AppendLine("<tr>");
                        currentVendor.HtmlData.AppendLine("<td>" + EscapeHtml(reader["GRN No"].ToString()) + "</td>");
                        currentVendor.HtmlData.AppendLine("<td>" + EscapeHtml(Convert.ToDateTime(reader["GRN Date"]).ToString("yyyy-MM-dd")) + "</td>");
                        currentVendor.HtmlData.AppendLine("<td>" + EscapeHtml(reader["Invoice No"].ToString()) + "</td>");
                        currentVendor.HtmlData.AppendLine("<td>" + EscapeHtml(Convert.ToDateTime(reader["Invoice Date"]).ToString("yyyy-MM-dd")) + "</td>");


                        currentVendor.HtmlData.AppendLine("<td>" + EscapeHtml(reader["Vendor Code"].ToString()) + "</td>");
                        currentVendor.HtmlData.AppendLine("<td>" + EscapeHtml(reader["Vendor Name"].ToString()) + "</td>");
               
                        currentVendor.HtmlData.AppendLine("<td>" + EscapeHtml(reader["ItemCode"].ToString()) + "</td>");
                        currentVendor.HtmlData.AppendLine("<td>" + EscapeHtml(reader["Dscription"].ToString()) + "</td>");
                        currentVendor.HtmlData.AppendLine("<td>" + EscapeHtml(Convert.ToDecimal(reader["Quantity"]).ToString("N2")) + "</td>");
                        currentVendor.HtmlData.AppendLine("<td>" + EscapeHtml(Convert.ToDecimal(reader["Price"]).ToString("N2")) + "</td>");
                        currentVendor.HtmlData.AppendLine("<td>" + EscapeHtml(Convert.ToDecimal(reader["LineTotal"]).ToString("N2")) + "</td>");



                        currentVendor.HtmlData.AppendLine("</tr>");

                        int docEntry = Convert.ToInt32(reader["DocEntry"]);

                        if (!currentVendor.DocEntries.Contains(docEntry))
                        {
                            currentVendor.DocEntries.Add(docEntry);
                        }
                    }

                    foreach (KeyValuePair<string, VendorGRPOData> vendor in vendors)
                    {
                        ProcessVendorGRPO(connection, vendor.Value);
                    }
                }
            }
            catch (Exception ex)
            {
                Log("GRPO process failed: " + ex.Message);
            }
        }

        private static void ProcessVendorGRPO(OdbcConnection connection, VendorGRPOData vendor)
        {
            try
            {
                vendor.HtmlData.AppendLine("</table>");

                SendEmail(vendor.ToMail, vendor.CcMail, vendor.CardCode, vendor.CardName, vendor.HtmlData.ToString());

                UpdateEmailStatus(connection, vendor.DocEntries);

                Log("Vendor " + vendor.CardCode + " - " + vendor.CardName + " processed successfully. Email sent to " + vendor.ToMail + ". GRPO count: " + vendor.DocEntries.Count);
            }
            catch (Exception ex)
            {
                Log("Vendor " + vendor.CardCode + " - " + vendor.CardName + " failed: " + ex.Message);
            }
        }

        private static void SendEmail(string toMail, string ccMail, string cardCode, string cardName, string htmlTable)
        {
            using (MailMessage mail = new MailMessage())
            using (SmtpClient smtp = new SmtpClient(SmtpServer, SmtpPort))
            {
                mail.From = new MailAddress(FromMail);
                mail.To.Add(toMail);

                if (!string.IsNullOrWhiteSpace(ccMail))
                {
                    mail.CC.Add(ccMail);
                }

                mail.Subject = "GRPO Report - " + cardCode + " - " + cardName;

                mail.IsBodyHtml = true;

                mail.Body = "<html><body style='font-family:Arial; font-size:12px;'>" +
              "<p>Dear Sir / Madam,</p>" +
              "<p>This is to inform you that the following materials are received. GRN details as follows:</p>" +
              "<br/>" +
              htmlTable +
              "<br/><br/>" +
              "<p>Thanks and Regards<br/>" +
              //"Finance Department.<br/><br/>" +
              "Amwin Machining Pvt Ltd.,<br/>" +
              //"Ground Floor, Sy.No. 41/2 &amp; 41/3,<br/>" +
              //"Karivobanahalli, Peenya 2nd Stage.<br/>" +
              //"GSTIN :- 29AARCA1318B1ZW</p>" +
              "</body></html>";

                smtp.EnableSsl = true;
                smtp.Credentials = new NetworkCredential(SmtpUser, SmtpPassword);
                smtp.Send(mail);
            }
        }

        private static void UpdateEmailStatus(OdbcConnection connection, List<int> docEntries)
        {
            foreach (int docEntry in docEntries)
            {
                string query = "UPDATE OPDN SET \"U_EmailSts\" = 'Yes' WHERE \"DocEntry\" = ?";

                using (OdbcCommand command = new OdbcCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@DocEntry", docEntry);
                    command.ExecuteNonQuery();
                }
            }
        }

        private static string EscapeHtml(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return "";
            }

            return value.Replace("&", "&amp;")
                        .Replace("<", "&lt;")
                        .Replace(">", "&gt;")
                        .Replace("\"", "&quot;")
                        .Replace("'", "&#39;");
        }

        private static void InitializeLog()
        {
            Directory.CreateDirectory(LogDirectory);

            File.AppendAllText(
                LogFilePath,
                Environment.NewLine +
                "===== GRPO Process Started: " +
                DateTime.Now +
                " =====" +
                Environment.NewLine);
        }

        private static void Log(string message)
        {
            Directory.CreateDirectory(LogDirectory);

            File.AppendAllText(
                LogFilePath,
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") +
                " - " +
                message +
                Environment.NewLine);
        }

        private class VendorGRPOData
        {
            public string CardCode { get; set; }
            public string CardName { get; set; }
            public string ToMail { get; set; }
            public string CcMail { get; set; }
            public StringBuilder HtmlData { get; set; }

            public List<int> DocEntries { get; set; } = new List<int>();
        }
    }
}
