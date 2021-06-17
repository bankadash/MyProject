using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
             Outlook.Application Application = new Microsoft.Office.Interop.Outlook.Application();
             Outlook.Folder root = Application.Session.DefaultStore.GetRootFolder() as Microsoft.Office.Interop.Outlook.Folder;
             EnumerateFolders(root);
        }

        // Uses recursion to enumerate Outlook subfolders.
        private void EnumerateFolders(Microsoft.Office.Interop.Outlook.Folder folder)
        {
             
            Microsoft.Office.Interop.Outlook.Folders childFolders = folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Microsoft.Office.Interop.Outlook.Folder childFolder in childFolders)
                {
                    // We only want Inbox folders - ignore Contacts and others
                    if (childFolder.FolderPath.Contains("SmartTrackerData"))
                    {
                        // Write the folder path.
                        Console.WriteLine(childFolder.FolderPath);
                        // Call EnumerateFolders using childFolder, to see if there are any sub-folders within this one
                        EnumerateFolders(childFolder);
                    }
                }
            }
            Console.WriteLine("Checking in " + folder.FolderPath);
            IterateMessages(folder);
        }

        static void IterateMessages(Microsoft.Office.Interop.Outlook.Folder folder)
        {
            // attachment extensions to save
            string[] extensionsArray = {".csv"};

            // Iterate through all items ("messages") in a folder
            var fi = folder.Items;
            if (fi != null)
            {

                try
                {
                    foreach (Object item in fi)
                    {

                        

                        Microsoft.Office.Interop.Outlook.MailItem mi = (Microsoft.Office.Interop.Outlook.MailItem)item;

                        if (mi.SentOn.ToShortDateString() == DateTime.Now.ToShortDateString())
                        {
                            var attachments = mi.Attachments;
                            if (attachments.Count != 0)
                            {

                                for (int i = 1; i <= mi.Attachments.Count; i++)
                                {
                                    var fn = mi.Attachments[i].FileName.ToLower();

                                    //check wither any of the strings in the extensionsArray are contained within the filename
                                    if (extensionsArray.Any(fn.Contains))
                                    {
                                        mi.Attachments[i].SaveAsFile(@"c:\SmartData\" + fn.ToString());
                                    }
                                }
                            }
                        }
                    }
                }
                catch (System.Exception e)
                {
                    //Console.WriteLine("An error occurred: '{0}'", e);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            // Create the Outlook application.
            Outlook.Application oApp = new Outlook.Application();
            // Create a new mail item.
            Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
            // Set HTMLBody. 
            //add the body of the email
            oMsg.HTMLBody = "Test body";
            //Add an attachment.
            String sDisplayName = "MyAttachment";
            int iPosition = (int)oMsg.Body.Length + 1;
            int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
            //now attached the file
            Outlook.Attachment oAttach = oMsg.Attachments.Add
                                         (@"C:\srikanta\xyz.csv", iAttachType, iPosition, sDisplayName);
            //Subject line
            oMsg.Subject = "test";
            // Add a recipient.
            Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
            // Change the recipient in the next line if necessary.
            Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add("srikanta.dash@accenture.com");
            oRecip.Resolve();
            // Send.
            oMsg.Send();
            // Clean up.
            oRecip = null;
            oRecips = null;
            oMsg = null;
            oApp = null;
        }
    }
}
