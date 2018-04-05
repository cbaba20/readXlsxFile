using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace ReadExcelAndSendMail
{
    class Program
    {
        public static string readExcelFile()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWb;
            Excel.Worksheet xlWs;
            Excel.Range range;
            string xlPath = @"C:\Users\Chandan\Desktop\DailyWorks.xlsx";

            int rowCount = 0;
            int columnCount = 0;
            string returningText = "";
            try
            {
                xlApp = new Excel.Application();
                xlWb = xlApp.Workbooks.Open(xlPath);
                xlWs = (Excel.Worksheet)xlWb.Worksheets.get_Item(1);

                range = xlWs.UsedRange;
                int lastRowNumber = range.Rows.Count;

                Excel.Range oRange = range.Cells[lastRowNumber, 1];
                Double loggedDate = double.Parse(Convert.ToString(oRange.Value2));
                DateTime changedDate = DateTime.FromOADate(loggedDate);
                if (changedDate.ToString("dd/MM/yyyy") == DateTime.Today.ToShortDateString())
                {
                    oRange = range.Cells[lastRowNumber, 2];
                    string task = oRange.Value2;

                    if (task!=null)
                    {
                        returningText =  "Today"+" "+ changedDate.ToString("dd/MM/yyyy")+" "+"I am doing" +" " +task +" "+"to my friends.";
                        //Console.WriteLine(changedDate.ToString("dd/MM/yyyy") + "-->" + task);
                    }
                }

                // if(range.Cells[i, 1] != null && range.Cells[i, 1].Value2==DateTime.Today.ToString("dd/MM/yyyy"))
                //for (int i = 1; i <= rw; i++)
                //{
                //    for (int j = 1; j <= cl; j++)
                //    {
                //        //new line
                //        if (j == 1)
                //            Console.Write("\r\n");

                //        //write the value to the console
                //        if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
                //            Console.Write(range.Cells[i, j].Value2.ToString() + "\t");
                //    }
                //}
                xlWb.Close();
                xlApp.Quit();
                Marshal.ReleaseComObject(range);
                Marshal.ReleaseComObject(xlWs);
                Marshal.ReleaseComObject(xlWb);
                Marshal.ReleaseComObject(xlApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                return returningText;

            }
            catch (Exception ex)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                throw ex;
            }
        }

        public static void sendMail()
        {
            Console.WriteLine("Mail To");
            MailAddress to = new MailAddress(Console.ReadLine());

            Console.WriteLine("Mail From");
            MailAddress from = new MailAddress(Console.ReadLine());

            MailMessage mail = new MailMessage(from, to);

            Console.WriteLine("Subject");
            mail.Subject = Console.ReadLine();

            Console.WriteLine("Your Message");
            mail.Body = readExcelFile();

            SmtpClient smtp = new SmtpClient();
            smtp.Host = "smtp.gmail.com";
            smtp.Port = 587;

            smtp.Credentials = new NetworkCredential(
                "moodler18@gmail.com", "Eclerx#123");
            smtp.EnableSsl = true;
            Console.WriteLine("Sending email...");
            smtp.Send(mail);
            Console.WriteLine("Email Sent.");
            System.Threading.Thread.Sleep(3000);
            Environment.Exit(0);
        }
        static void Main(string[] args)
        {
            readExcelFile();
            sendMail();
        }
    }
}