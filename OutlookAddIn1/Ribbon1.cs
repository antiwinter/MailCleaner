using System;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Windows.Forms;
using System.Threading;


namespace OutlookAddIn1
{
    public partial class Ribbon1
    {

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            this.OlApp = new Outlook.Application();
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
                // Break recursive
                if (MainForm.StopState == 1)
                {
                    return;
                    MainForm.StopState = 2;
                }

                if (m is Outlook.MailItem)
                {
                    Outlook.MailItem _m = (Outlook.MailItem)m;

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
                    BadObjects++;

#if false
                    // Delete bad obj and return
                    try
                    {
                        BadFolder = TrashFolder.Folders["CLEAN_BadObjs"];
                    }
                    catch (System.Exception e1)
                    {
                        BadFolder = TrashFolder.Folders.Add("CLEAN_BadObjs");
                    }

                    _m.Move(BadFolder);
#endif
                    return;
                }

                // Item is identical
                Debug.WriteLine("DupMail/sub: " + _m.CreationTime + _m.Subject);
                ItemsFound++;

                // Delete this item and return
                try
                {
                    DupFolder = TrashFolder.Folders["CLEAN_Duplicates"];
                }
                catch (System.Exception e2)
                {
                    DupFolder = TrashFolder.Folders.Add("CLEAN_Duplicates");
                }

                _m.Move(DupFolder);
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
            button1.Enabled = false;
            CurrentIndex = 1;
            ItemsFound = BadObjects = TotalIndex
                = Duplicate_Buffer_Cursor = 0;

            // MainFolder = OlApp.Session.PickFolder();
            MainFolder = OlApp.ActiveExplorer().CurrentFolder;
            TrashFolder = OlApp.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems);
            Debug.WriteLine("current folder: " + MainFolder.Name);
            Debug.WriteLine("delete folder: " + TrashFolder.Name);


            CalcTotalItems(MainFolder);
            Debug.WriteLine("Items to be parsed: " + TotalIndex);

            MainForm = new Form1();
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

            // The TotalIndex may be a little bigger than the parsed items
            // set CurrentIndex to TotalIndex to ensure the progressbar reach
            // the end
            CurrentIndex = TotalIndex;
            Thread.Sleep(500);

            MainForm.StopState = 2;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (CurrentIndex <= MainForm.progressBar1.Maximum)
            {
                MainForm.label3.Text = CurrentIndex + " of " + TotalIndex + " checked";
                if(ItemsFound > 0)
                    MainForm.label4.Text = ItemsFound + " duplicates";
//                if (BadObjects > 0)
//                    MainForm.label5.Text = BadObjects + " bad objects";

                MainForm.progressBar1.Value = CurrentIndex;
                MainForm.progressBar1.PerformStep();
            }

            if( MainForm.StopState == 2)
            {
                MainForm.progressBar1.Visible = false;
                MainForm.label4.Text = ItemsFound + " duplicates";
//                    MainForm.label5.Text = BadObjects + " bad objects";
                MainForm.linkLabel1.Text = (ItemsFound > 0 /* || BadObjects > 0 */) ?
                    "Cleaning finished" :
                    "Okay, no trash found";
                timer1.Stop();
                button1.Enabled = true;
            }
        }

        private Outlook.Application OlApp;
        private Form1 MainForm;
        private Outlook.MAPIFolder MainFolder, TrashFolder, BadFolder, DupFolder;
        private Outlook.MailItem[] Duplicate_Buffer;
        private int CurrentIndex, TotalIndex, ItemsFound, BadObjects,
            MaxComapre = 3, Duplicate_Buffer_Cursor;

    }
}
