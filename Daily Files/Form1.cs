//Created by Morgan Thompson 
//Updated: 5/13/2014 
//Contact: morgan@artive.co

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
using Microsoft.Office;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Reflection;
using Access = Microsoft.Office.Interop.Access;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop;

namespace Daily_Files
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            logBox.ScrollBars = ScrollBars.Vertical;
            monthCalendar1.MaxDate = DateTime.Today;
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

        private string getEXDPath (string day)
        {   
            DateTime today = DateTime.Today;
            if (day == "monday")
            {
                return exdFile(today.Day, today.Month, today.Year);
            }
            if (day == "sunday")
            {
                DateTime yesterday = today.AddDays(-1); // Sunday
                return exdFile(yesterday.Day, yesterday.Month, yesterday.Year);
            }
            if (day == "saturday")
            {
                DateTime minusTwoDays = today.AddDays(-2); // Saturday
                return exdFile(minusTwoDays.Day, minusTwoDays.Month, minusTwoDays.Year);
            }
            else
            {
                string nil = "";
                return nil;
            }
        }

        private void email()
        {
            DateTime today = DateTime.Today;
            logBox.AppendText(Environment.NewLine);
            logBox.AppendText("Checking for EXD files...");
            if (((File.Exists(getEXDPath("monday"))) == true) && (((today.DayOfWeek == DayOfWeek.Tuesday) || (today.DayOfWeek == DayOfWeek.Wednesday) || (today.DayOfWeek == DayOfWeek.Thursday)
            || (today.DayOfWeek == DayOfWeek.Friday)))) //If it's not monday, just check if there is an EXD file for that day
            {
                sendEmail();
            }

            else if ((today.DayOfWeek == DayOfWeek.Monday) && ((File.Exists(getEXDPath("monday"))) || ((File.Exists(getEXDPath("sunday")))) || ((File.Exists(getEXDPath("saturday"))))))
            // Checks to see if there are EXD files saved for Monday, Sunday, or Saturday
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
            DateTime today = DateTime.Today;
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
                if (today.DayOfWeek == DayOfWeek.Monday)
                {
                    if ((File.Exists(getEXDPath("monday")))) // Checks to see if there is an EXD file for Monday 
                    { Outlook.Attachment aAttach = oMsg.Attachments.Add(getEXDPath("monday"), iAttachType, iPosition, sDisplayName); }

                    if ((File.Exists(getEXDPath("sunday")))) // Checks to see if there is an EXD file for Sunday 
                    { Outlook.Attachment bAttach = oMsg.Attachments.Add(getEXDPath("sunday"), iAttachType, iPosition, sDisplayName); }

                    if ((File.Exists(getEXDPath("saturday")))) // Checks to see if there is an EXD file for Saturday
                    { Outlook.Attachment cAttach = oMsg.Attachments.Add(getEXDPath("saturday"), iAttachType, iPosition, sDisplayName); }
                }
                else
                {
                    if ((File.Exists(exdFile(today.Day, today.Month, today.Year))))
                    {
                        Outlook.Attachment aAttach = oMsg.Attachments.Add(exdFile(today.Day, today.Month, today.Year), iAttachType, iPosition, sDisplayName);
                    }
                }

                //Subject line
                oMsg.Subject = "EXD Daily Files";
                // Add a recipient.
                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                // Change the recipient in the next line if necessary.
                Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add("aschuber@ccs.ua.edu");
                oRecip.Resolve();
                //Send.
                oMsg.Send();
                // Clean up.
                oRecip = null;
                oRecips = null;
                oMsg = null;
                oApp = null;
            }//end of try block
            catch (Exception ex)
            {
                logBox.AppendText(Environment.NewLine);
                logBox.AppendText(ex.ToString());
            }//end of catch
            System.Diagnostics.ProcessStartInfo p = new System.Diagnostics.ProcessStartInfo(@"J:\Academic Outreach\Banner Files\BannerData.mdb");
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            proc.Close();
            return; 
        }
 
        private string exdFile(int day, int month, int year)
        {
            string file = @"J:\Academic Outreach\Banner Files\EXD\EXD_";
            int zero = 0;
           
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
            Microsoft.Office.Interop.Access.Application oAccess = new Microsoft.Office.Interop.Access.Application();
            oAccess.Visible = false;
            oAccess.OpenCurrentDatabase(@"J:\Academic Outreach\Banner Files\BannerData.mdb", false, "");
            string macro = "Automatic_Output";
            oAccess.DoCmd.RunMacro(macro);
        }

        private void killAccess()
        {
            foreach (Process p in System.Diagnostics.Process.GetProcessesByName("MSACCESS"))
            {
                try
                {
                    p.Kill();
                    p.WaitForExit(); // possibly with a timeout
                }
                catch (Win32Exception winException)
                {
                    // process was terminating or can't be terminated - deal with it
                }
                catch (InvalidOperationException invalidException)
                {
                    // process has already exited - might be able to let this one go
                }
            }
        }

        private string WeekdayFilePath(int day, int month, int year)
        {
            string filePath = @"\\NEWDISTED2\ftproot\archive\archive_ao_regfile";
            int zero = 0;
            string code = @"0946.txt";
            string fileDate = year.ToString();
            if (month < 10) 
            { 
                fileDate += zero.ToString();
                fileDate += month.ToString();
            }
            else 
            { 
                fileDate += month.ToString(); 
            }
            if (day < 10)
            { 
                fileDate += zero.ToString();
                fileDate += day.ToString();
            }
            else 
            {
                fileDate += day.ToString();
            }
            fileDate += code;
            filePath += fileDate;
            return filePath;
        }

        private string dateRangeFilePath( int year, int day, int month)
        {
            string filePath = @"\\NEWDISTED2\ftproot\archive\archive_ao_regfile";
            int zero = 0;
            string code = @"0946.txt";
            string fileDate = year.ToString();
            if (month < 10)
            {
                fileDate += zero.ToString();
                fileDate += month.ToString();
            }
            else
            {
                fileDate += month.ToString();
            }
            if (day < 10)
            {
                fileDate += zero.ToString();
                fileDate += day.ToString();
            }
            else
            {
                fileDate += day.ToString();
            }
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
            logBox.AppendText("Creating EXD file for: ");
            logBox.AppendText(DateTime.Today.Month.ToString());
            logBox.AppendText("/");
            logBox.AppendText(DateTime.Today.Day.ToString());
            logBox.AppendText("/");
            logBox.AppendText(DateTime.Today.Year.ToString());
            openAccess();
            email();
            logBox.AppendText(Environment.NewLine);
            logBox.AppendText("Done!");
            killAccess();
        }

        private void mondayButton_Click_1(object sender, EventArgs e)
        {
            DateTime today = DateTime.Today;

                if (File.Exists(WeekdayFilePath(today.Day,today.Month, today.Year)))
                {
                    //This is Monday's file
                    File.Copy(WeekdayFilePath(today.Day,today.Month, today.Year), @"J:\Academic Outreach\Banner Files\ao_regfile.txt", true);
                    logBox.AppendText(Environment.NewLine);
                    logBox.AppendText("Creating EXD file for today...");
                    openAccess();
                }
                else
                {
                    logBox.AppendText(Environment.NewLine);
                    logBox.AppendText("The archive file for today was not found...");
                }
                
                DateTime yesterday = today.AddDays(-1); // Yesterday
                if (File.Exists(WeekdayFilePath(yesterday.Day, yesterday.Month, yesterday.Year)))
                {
                    File.Copy(WeekdayFilePath(yesterday.Day, yesterday.Month, yesterday.Year), @"J:\Academic Outreach\Banner Files\ao_regfile.txt", true);
                    logBox.AppendText(Environment.NewLine);
                    logBox.AppendText("Creating EXD file for Sunday...");
                    openAccess();
                }
                else
                {
                    logBox.AppendText(Environment.NewLine);
                    logBox.AppendText("The archive file for Sunday was not found...");
                }
                
                DateTime minusTwoDays = today.AddDays(-2); // Saturday
                if (File.Exists(WeekdayFilePath(minusTwoDays.Day, minusTwoDays.Month, minusTwoDays.Year)))
                {
                    File.Copy(WeekdayFilePath(minusTwoDays.Day, minusTwoDays.Month, minusTwoDays.Year), @"J:\Academic Outreach\Banner Files\ao_regfile.txt", true);
                    logBox.AppendText(Environment.NewLine);
                    logBox.AppendText("Creating EXD file for Saturday...");
                    openAccess();
                }
                else
                {
                    logBox.AppendText(Environment.NewLine);
                    logBox.AppendText("The archive file for Saturday was not found...");
                }
            //Saved all of the EXD files... sending email!
            email();
            logBox.AppendText(Environment.NewLine);
            logBox.AppendText("Done!");
            killAccess();
        }  

        private string exdFileName(int month, int day, int year)
        {
            int zero = 0;
            string file = "EXD_";
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

        private void button1_Click(object sender, EventArgs e)
        {
            string folder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            folder += @"\EXD Files Generated on ";
            folder += DateTime.Today.Month.ToString();
            folder += ".";
            folder += DateTime.Today.Day.ToString();
            folder += ".";
            folder += DateTime.Today.Year.ToString();
            System.IO.Directory.CreateDirectory(folder);
            folder += @"\";
            string interFolder = folder;

            List<DateTime> allDates = new List<DateTime>();
            for (DateTime date = monthCalendar1.SelectionStart; date <= monthCalendar1.SelectionEnd; date = date.AddDays(1))
                allDates.Add(date);
            for (int i = 0; i < allDates.Count; i++)
            {
                if (File.Exists(dateRangeFilePath(allDates[i].Year, allDates[i].Day, allDates[i].Month)))
                {
                    File.Copy(dateRangeFilePath(allDates[i].Year, allDates[i].Day, allDates[i].Month), @"J:\Academic Outreach\Banner Files\ao_regfile.txt", true);
                    
                    if (File.Exists(exdFile(allDates[i].Day, allDates[i].Month, allDates[i].Year)))
                    {
                        logBox.AppendText(Environment.NewLine);
                        logBox.AppendText("The EXD file for ");
                        logBox.AppendText(allDates[i].Month.ToString());
                        logBox.AppendText("/");
                        logBox.AppendText(allDates[i].Day.ToString());
                        logBox.AppendText("/");
                        logBox.AppendText(allDates[i].Year.ToString());
                        logBox.AppendText(" already exists! ");
                    }
                    else 
                    {
                        logBox.AppendText(Environment.NewLine);
                        logBox.AppendText("Creating EXD file for: ");
                        logBox.AppendText(allDates[i].Month.ToString());
                        logBox.AppendText("/");
                        logBox.AppendText(allDates[i].Day.ToString());
                        logBox.AppendText("/");
                        logBox.AppendText(allDates[i].Year.ToString());
                        openAccess(); 
                    }
                    
                    if (File.Exists(exdFile(allDates[i].Day, allDates[i].Month, allDates[i].Year)))
                    {
                        folder += (exdFileName(allDates[i].Month, allDates[i].Day, allDates[i].Year));
                        File.Copy(exdFile(allDates[i].Day, allDates[i].Month, allDates[i].Year), folder, true);
                        folder = interFolder;
                    }
                }
                else
                {
                    logBox.AppendText(Environment.NewLine);
                    logBox.AppendText("There is no archive file for: ");
                    logBox.AppendText(allDates[i].Month.ToString());
                    logBox.AppendText("/");
                    logBox.AppendText(allDates[i].Day.ToString());
                    logBox.AppendText("/");
                    logBox.AppendText(allDates[i].Year.ToString());
                }
            }
             
            Process.Start(folder); //Open up the new folder that contains newly created EXD files
            
            logBox.AppendText(Environment.NewLine);
            logBox.AppendText("Done!");
            killAccess();
        }
    }
}

