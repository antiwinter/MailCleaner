using System;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Windows.Forms;


namespace OutlookAddIn1
{
    public partial class Ribbon1
    {

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            this.OlApp = new Outlook.Application();
            this.CurrentIndex = 1;
            this.TotalIndex = 0;
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
                    // Call EnumerateFolders using childFolder.
                    EnumerateFolders(childFolder);
                }
            }

            // Write the folder path.
            Debug.WriteLine("Checking: " + folder.FolderPath);
            foreach (Object m in folder.Items)
            {
                if (m is Outlook.MailItem)
                {
                    Outlook.MailItem _m = (Outlook.MailItem)m;
                    if (CurrentIndex > 11620)
                    {
#if false
                        try
                        {
                            Debug.WriteLine("subject: " + _m.Subject);
                            Debug.WriteLine("create time: " + _m.CreationTime);
                            Debug.WriteLine("receiv time: " + _m.ReceivedTime);
                            Debug.WriteLine("sender id: " + _m.Sender.ID);
                            Debug.WriteLine("sender addr: " + _m.Sender.Address);
                            Debug.WriteLine("///sender addr: " + _m.SenderEmailAddress);
                            Debug.WriteLine("///sender type: " + _m.SenderEmailType);
                            Debug.WriteLine("///sender name: " + _m.SenderName);
                            Debug.WriteLine("recipe id: " + _m.Sender.ID);
                            Debug.WriteLine("recipe addr: " + _m.Sender.Address);
                            Debug.WriteLine("size: " + _m.Size);
                            Debug.WriteLine("to: " + _m.To);
                            Debug.WriteLine("cc: " + _m.CC);
                            Debug.WriteLine("attachment count: " + _m.Attachments.Count);
                            Debug.WriteLine(" ");
                        }

                        catch (System.Exception e)
                        {
                            Debug.WriteLine("bad obj: " + _m.Subject);
                        }
#endif
                    }

                    // Parse the item with function: find duplicates
                    FindDuplicates(_m);
                    CurrentIndex++;
                }
            }
        }

        private void FindDuplicates(Outlook.MailItem _m)
        {
            for(int i = 0; i < MaxComapre; i++)
            {
                if (Duplicate_Buffer[i] == null)
                    continue;

                try {
                    // Check creation time
                    if (DateTime.Compare(_m.ReceivedTime, Duplicate_Buffer[i].ReceivedTime) != 0)
                        continue;

                    // Check subject
                    if (String.Compare(_m.Subject, Duplicate_Buffer[i].Subject) != 0)
                        continue;

                    // Check sender
                    if (String.Compare(_m.SenderEmailAddress, Duplicate_Buffer[i].SenderEmailAddress) != 0)
                        continue;

                    // Check TO
                    if (String.Compare(_m.To, Duplicate_Buffer[i].To) != 0)
                        continue;

                    // Check CC
                    if (String.Compare(_m.CC, Duplicate_Buffer[i].CC) != 0)
                        continue;

                    // Check Attachments count
                    if (_m.Attachments.Count != Duplicate_Buffer[i].Attachments.Count)
                        continue;

                    // Check body
                    if (String.Compare(_m.Body, Duplicate_Buffer[i].Body) != 0)
                        continue;
                }

                catch (System.Exception e)
                {
                    Debug.WriteLine("bad obj: " + _m.Subject);
                    Debug.WriteLine(e.Message);

                    // Delete bad obj and return
                    return;
                }

                // Item is identical
                Debug.WriteLine("DupMail/sub: " + _m.CreationTime + _m.Subject);
                ItemsFound++;

                // Delete this item and return
                return;
            }

            Duplicate_Buffer[Duplicate_Buffer_Cursor] = _m;
            Duplicate_Buffer_Cursor = (Duplicate_Buffer_Cursor + 1) % 3;
        }

        private void CalcTotalItems(Outlook.MAPIFolder folder)
        {
            Outlook.Folders childFolders =
                folder.Folders;

            TotalIndex += folder.Items.Count;

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
            MainForm = new Form1();
            // MainFolder = OlApp.Session.PickFolder();
            MainFolder = OlApp.ActiveExplorer().CurrentFolder;
            Debug.WriteLine("current folder: " + MainFolder.Name);

            CalcTotalItems(MainFolder);

            Debug.WriteLine("Items to be parsed: " + TotalIndex);
            MainForm.progressBar1.Minimum = 1;
            MainForm.progressBar1.Maximum = TotalIndex;
            MainForm.progressBar1.Step = 1;

            Duplicate_Buffer = new Outlook.MailItem[MaxComapre];

            backgroundWorker1.RunWorkerAsync();
            timer1.Start();
            MainForm.Show();
        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            EnumerateFolders(MainFolder);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (CurrentIndex <= MainForm.progressBar1.Maximum)
            {
                MainForm.progressBar1.Value = CurrentIndex;
                MainForm.label1.Text = "Items found: " + ItemsFound;
                MainForm.progressBar1.PerformStep();
            }
        }

        private Outlook.Application OlApp;
        private Form1 MainForm;
        private Outlook.MAPIFolder MainFolder;
        private Outlook.MailItem[] Duplicate_Buffer;
        private int CurrentIndex, TotalIndex, ItemsFound,
            MaxComapre = 3, Duplicate_Buffer_Cursor = 0;

    }
}
