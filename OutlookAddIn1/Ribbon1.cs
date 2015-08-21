using System;
using System.Threading;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Windows.Forms;

namespace OutlookAddIn1
{
    public partial class Ribbon1
    {
        private Outlook.Application OlApp;
        private int Count, Total;
        Form1 progressForm;
        Outlook.MAPIFolder folder;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            this.OlApp = new Outlook.Application();
            this.Count = 1;
            this.Total = 0;
        }

        private void parse_Mails()
        {

#if false
            if (folder != null)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("Folder EntryID:");
                sb.AppendLine(folder.EntryID);
                sb.AppendLine();
                sb.AppendLine("Folder StoreID:");
                sb.AppendLine(folder.StoreID);
                sb.AppendLine();
                sb.AppendLine("Unread Item Count: "
                    + folder.UnReadItemCount);
                sb.AppendLine("Default MessageClass: "
                    + folder.DefaultMessageClass);
                sb.AppendLine("Current View: "
                    + folder.CurrentView.Name);
                sb.AppendLine("Folder Path: "
                    + folder.FolderPath);
                MessageBox.Show(sb.ToString(),
                    "Folder Information",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
#endif

        }

        // Uses recursion to enumerate Outlook subfolders.
        private void EnumerateFolders(Outlook.MAPIFolder folder)
        {
            Outlook.Folders childFolders =
                folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Outlook.Folder childFolder in childFolders)
                {
                    // Write the folder path.
                    Debug.WriteLine(childFolder.FolderPath);
                    // Call EnumerateFolders using childFolder.
                    EnumerateFolders(childFolder);
                }
            }

            foreach (Object m in folder.Items)
            {
                if (m is Outlook.MailItem)
                {
                    Outlook.MailItem _m = (Outlook.MailItem)m;
                    Debug.WriteLine(_m.Subject);
                    Count++;
                }
            }
        }

        private void CalcTotalItems(Outlook.MAPIFolder folder)
        {
            Outlook.Folders childFolders =
                folder.Folders;

            Total += folder.Items.Count;

            if (childFolders.Count > 0)
            {
                foreach (Outlook.Folder childFolder in childFolders)
                {
                    CalcTotalItems(childFolder);
                }
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            progressForm = new Form1();
            folder = OlApp.Session.PickFolder();

            CalcTotalItems(folder);
            MessageBox.Show("all: " + Total,
            "Count Information",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);

            progressForm.progressBar1.Minimum = 1;
            progressForm.progressBar1.Maximum = folder.Items.Count;
            progressForm.progressBar1.Step = 1;
            backgroundWorker1.RunWorkerAsync();
            timer1.Start();
            progressForm.Show();
        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            EnumerateFolders(folder);
            MessageBox.Show("Count: " + Count,
            "Count Information",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (Count <= progressForm.progressBar1.Maximum)
            {
                progressForm.progressBar1.Value = Count;
                progressForm.label1.Text = "Items found: " + Count;
                progressForm.progressBar1.PerformStep();
            }
        }
    }
}
