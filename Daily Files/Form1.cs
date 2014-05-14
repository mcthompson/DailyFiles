//Created by Morgan Thompson 
//Updated: 5/13/2014 
//Contact: morganchristopherthompson@gmail.com 

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Daily_Files
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            dateTo.MaxDate = DateTime.Today;
            dateFrom.MaxDate = DateTime.Today;
            DayOfWeek today = DateTime.Today.DayOfWeek;
            if (today == DayOfWeek.Monday)
            {
                tfButton.Enabled = false;
            }
            else
            {
                mondayButton.Enabled = false;
            }
        }

        public void email()
        {
            DayOfWeek today = DateTime.Today.DayOfWeek;
            logBox.AppendText(Environment.NewLine);
            logBox.AppendText("Checking for EXD files...");
            if(((File.Exists(exdFile(0))) == true) && (((today == DayOfWeek.Tuesday) || (today == DayOfWeek.Wednesday) || (today == DayOfWeek.Thursday)
            || (today == DayOfWeek.Friday))))
            {
                sendEmail();
            }

            else if ((today == DayOfWeek.Monday) && (((File.Exists(exdFile(0)))) || ((File.Exists(exdFile(1)))) || ((File.Exists(exdFile(2))))))
            // Checks to see if there are EXD files saved for today, 1 day prior, 2 days prior
            {
                sendEmail();
            }
            else
            {
                logBox.AppendText(Environment.NewLine);
                logBox.AppendText("No EXD files found!");
                logBox.AppendText(Environment.NewLine);
                logBox.AppendText("No email will be sent.");
                return;
            }
        }

        private void sendEmail()
        {
            DayOfWeek today = DateTime.Today.DayOfWeek;
            Process.Start(@"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Office Outlook 2007.lnk");
            Task.Delay(10000);
            logBox.AppendText(Environment.NewLine);
            logBox.AppendText("Sending email...");
            try
            {
                // Create the Outlook application.
                Outlook.Application oApp = new Outlook.Application();
                // Create a new mail item.
                Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                // Set HTMLBody. 
                //add the body of the email
                oMsg.HTMLBody = "Please see attached.";
                //Add an attachment.
                String sDisplayName = "MyAttachment";
                int iPosition = (int)oMsg.Body.Length + 1;
                int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
                //now attached the file
                if (today == DayOfWeek.Monday)
                {
                    if ((File.Exists(exdFile(0)))) // Checks to see if there is an EXD file for Monday 
                    { Outlook.Attachment aAttach = oMsg.Attachments.Add(exdFile(0), iAttachType, iPosition, sDisplayName); }

                    if ((File.Exists(exdFile(1)))) // Checks to see if there is an EXD file for Sunday 
                    { Outlook.Attachment bAttach = oMsg.Attachments.Add(exdFile(1), iAttachType, iPosition, sDisplayName); }

                    if ((File.Exists(exdFile(2)))) // Checks to see if there is an EXD file for Saturday
                    { Outlook.Attachment cAttach = oMsg.Attachments.Add(exdFile(2), iAttachType, iPosition, sDisplayName); }
                }
                else
                {
                    if ((File.Exists(exdFile(0))))
                    {
                        Outlook.Attachment aAttach = oMsg.Attachments.Add(exdFile(0), iAttachType, iPosition, sDisplayName);
                    }
                }

                //Subject line
                oMsg.Subject = "EXD Daily Files";
                // Add a recipient.
                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                // Change the recipient in the next line if necessary.
                Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add("morganchristopherthompson@gmail.com");
                oRecip.Resolve();
                // Send.
                //oMsg.Send();
                // Clean up.
                oRecip = null;
                oRecips = null;
                oMsg = null;
                oApp = null;
            }//end of try block
            catch (Exception ex)
            {
            }//end of catch
        }
 
        private string exdFile(int value)
        {
            string file = @"J:\Academic Outreach\Banner Files\EXD\EXD_";
            int zero = 0;
            int month = DateTime.Today.Month;
            int day = DateTime.Today.Day - value;
            int year = DateTime.Today.Year;
            //File path includes a zero before month and day, but DateTime day/month does not contain zero. 
            //This will add zero before month/day if needed. 
            if (month <= 10) 
            { 
                file += zero.ToString();
                file += month.ToString();
            }
            else { file += month.ToString(); }
            if (day <= 10) 
            { 
                file += zero.ToString();
                file += day.ToString();
            }
            else { file += day.ToString(); }
            file += year.ToString();
            file += ".xls";
            return file;
        }

        private void openAccess()
        {
            System.Diagnostics.ProcessStartInfo p = new System.Diagnostics.ProcessStartInfo(@"J:\Academic Outreach\Banner Files\BannerData.mdb");
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            logBox.AppendText(Environment.NewLine);
            logBox.AppendText("Opening the Access Database...");
            proc.StartInfo = p;
            proc.Start();
            //Running the Access Database to generate EXD Excel spreadsheet
            proc.WaitForExit();
        }

        private string filePath(int value)
        {
            string filePath = @"\\NEWDISTED2\ftproot\archive\archive_ao_regfile";
            int year = DateTime.Today.Year;
            int zero = 0;
            int month = DateTime.Today.Month;
            int day = DateTime.Today.Day - value;
            string code = @"0946.txt";
            string fileDate = year.ToString();
            if (month <= 10) { fileDate += zero.ToString(); }
            else { fileDate += month.ToString(); }
            if (day <= 10) { fileDate += zero.ToString(); }
            else { fileDate += day.ToString(); }
            fileDate += code;
            filePath += fileDate;
            return filePath;
        }

        private void tfButton_Click(object sender, EventArgs e)
        {
            DayOfWeek today = DateTime.Today.DayOfWeek;
            Console.WriteLine(today);
            //Copying the ao_regfile to Banner Files
            System.IO.File.Copy(@"\\NEWDISTED2\ftproot\ao_regfile.txt", @"J:\Academic Outreach\Banner Files\ao_regfile.txt", true);
            logBox.AppendText(Environment.NewLine);
            logBox.AppendText("Copying ao_regfile to Banner Files...");
            openAccess();
            email();
            logBox.AppendText(Environment.NewLine);
            logBox.AppendText("Done!");
        }

        private void mondayButton_Click_1(object sender, EventArgs e)
        {
            //Copying the ao_regfile from Archive to Banner Files
            if (File.Exists(filePath(0)))
            {
                File.Copy(filePath(0), @"J:\Academic Outreach\Banner Files\ao_regfile.txt", true);  //This is Monday's file
                logBox.AppendText(Environment.NewLine);
                logBox.AppendText("Copying today's file...");
                openAccess();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Error! The file: " + filePath(0) + " could not be found!");
                return;
            }

            if (File.Exists(filePath(1)))
            {
                File.Copy(filePath(1), @"J:\Academic Outreach\Banner Files\ao_regfile.txt", true);  //This is Sunday's file
                logBox.AppendText(Environment.NewLine);
                logBox.AppendText("Copying Sunday's file...");
                openAccess();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Error! The file: " + filePath(1) + " could not be found!");
                return;
            }

            if (File.Exists(filePath(2)))
            {
                File.Copy(filePath(2), @"J:\Academic Outreach\Banner Files\ao_regfile.txt", true);  //This is Saturday's file
                logBox.AppendText(Environment.NewLine);
                logBox.AppendText("Copying Saturday's file...");
                openAccess();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Error! The file: " + filePath(2) + " could not be found!");
                return;
            }

            //Now we've saved all of the EXD files... let's open that email!
            email();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int fromDate = dateFrom.Value.Day;
            int toDate = dateTo.Value.Day;
            int days = toDate - fromDate;
            Console.WriteLine(days);
           
            if (dateTo.Value <= dateFrom.Value) { System.Windows.Forms.MessageBox.Show("Error: The 'To' date is before or equal to the 'From' date. Please fix this to continue."); }
            else
            {
                for (int i = 0; i < days; i++)
                {
                    if (File.Exists(filePath(i))) 
                    { 
                    File.Copy(filePath(i), @"J:\Academic Outreach\Banner Files\ao_regfile.txt", true);
                    openAccess();
                    }
                    }
            }
        }
    }
}

