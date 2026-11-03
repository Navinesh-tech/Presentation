using System;
using System.Configuration;
using System.Data.Odbc;

namespace EmailReport_AMWIN
{
    internal class Program
    {
        public static string HanaDriver = ConfigurationManager.AppSettings["HanaDriver"];
        public static string HanaServer = ConfigurationManager.AppSettings["HanaServer"];
        public static string HanaUser = ConfigurationManager.AppSettings["HanaUser"];
        public static string HanaPassword = ConfigurationManager.AppSettings["HanaPassword"];
        public static string HanaDatabase = ConfigurationManager.AppSettings["HanaDatabase"];

        public static string ConnectionString = "Driver=" + HanaDriver + ";ServerNode=" + HanaServer + ";UID=" + HanaUser + ";PWD=" + HanaPassword + ";CurrentSchema=" + HanaDatabase + ";";

        static void Main(string[] args)
        {
            try
            {
                using (OdbcConnection connection = new OdbcConnection(ConnectionString))
                {
                    connection.Open();

                    PO.ProcessPurchaseOrders(connection);
                    AR_Invoice.ProcessARInvoices(connection);
                    GRPO.ProcessGRPO(connection);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Application Error: " + ex.Message);
            }

            Console.WriteLine("Process completed.");
            Environment.Exit(0);
        }
    }
}
