using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Deployment.Application;
using Microsoft.Toolkit.Uwp.Notifications;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ExcelFileModify
{
    public partial class Form1 : Form
    {
        //private System.Timers.Timer checkForFile;

        [DllImport("user32.dll")]
        internal static extern IntPtr SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        internal static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private static System.Timers.Timer timer = new System.Timers.Timer();
        public Form1()
        {
            InitializeComponent();

            this.FormClosing += Form1_FormClosing;
            ToastNotificationManagerCompat.OnActivated += OnActivated;
            timer.AutoReset = true;
            
            timer.Elapsed += delegate (object sender, System.Timers.ElapsedEventArgs e)
            {
                //remove the StatusMessage text using a dispatcher, because timer operates in another thread
                this.BeginInvoke(new Action(() =>
                {
                    statusBar.Text = "";
                }));
            };
        }

        private void OnActivated(ToastNotificationActivatedEventArgsCompat e)
        {
            if (e is ToastNotificationActivatedEventArgsCompat toastActivationArgs)
            {
                // Obtain the arguments from the notification
                ToastArguments args = ToastArguments.Parse(toastActivationArgs.Argument);

                // Obtain any user input (text boxes, menu selections) from the notification
                //ValueSet userInput = toastActivationArgs.UserInput;

                // TODO: Show the corresponding content

                if (this.WindowState == FormWindowState.Minimized)
                {
                    this.WindowState = FormWindowState.Normal;
                }
                ShowWindow(this.Handle, 5);
                this.Activate(); 
                this.TopMost = true; // important
                this.TopMost = false; // important
                this.Focus();            
            }
        }

        private async void importData_click(object sender, EventArgs e)
        {
            if (progressBar1.Visible == false)
            {             
                string path = selectFile();
                if (path.Length > 0)
                {
                    this.Cursor = Cursors.WaitCursor;
                    progressBar1.Visible = true;
                    progressBar1.Style = ProgressBarStyle.Marquee;
                    ShowStatusMessage("Importing data...", 0);
                    try
                    {
                        await Task.Run(() => EditExcelFile.ImportExcelData(path));
                        ShowStatusMessage("Data imported successfully.", 5);
                        if (!this.ContainsFocus)
                            new ToastContentBuilder()
                                .AddText("Finished")
                                .AddText("Importing is done")
                                .Show();
                    }
                    catch (Exception ex)
                    {
                        ShowStatusMessage("Coulnd't import", 10);
                        if (ex is ArgumentNullException || ex is SQLiteException || ex is ArgumentException)
                            showError(ex.Message);
                        else if (ex is MissingFieldException)
                            showWarning(ex.Message);
                        else
                        {
                            showError(ex.Message);
                            //showError("Something went wrong the app will close now");
                            Application.Exit();                     
                        }
                    }
                }
                
                progressBar1.Visible = false;
                this.Cursor = Cursors.Default;
            }
        }
        private async void modify_excel_file_Click(object sender, EventArgs e)
        {
            Process currentProcess = Process.GetCurrentProcess();
            IntPtr hWnd = currentProcess.MainWindowHandle;

            if (progressBar1.Visible == false)
            {             
                string path = selectFile();
                //string dataPath = Environment.CurrentDirectory;
                string dataPath = ApplicationDeployment.CurrentDeployment.DataDirectory;
                dataPath += @"\tmpFiles\";
                
                if (path.Length > 0 && Directory.Exists(dataPath))
                {
                    this.Cursor = Cursors.WaitCursor;
                    progressBar1.Visible = true;
                    progressBar1.Style = ProgressBarStyle.Marquee;
                    
                    ShowStatusMessage("Modifying excel file...", 0);

                    try
                    {
                        List<string> modifiedFilesPath = await Task.Run(() => EditExcelFile.AddZones(path, inParts_chb.Checked));
                        progressBar1.Visible = false;
                        this.Cursor = Cursors.Default;
                        
                        if (!this.IsAccessible && !this.Focused)
                        {
                            if ("missing.txt" == Path.GetFileName(modifiedFilesPath[modifiedFilesPath.Count - 1]))
                            {
                                ShowStatusMessage("Some streets' data are missing. Check missing.txt file", 10);
                                if(!this.ContainsFocus)
                                    new ToastContentBuilder()
                                        .AddText("Finished")
                                        .AddText("Your excel file is ready")
                                        .AddText("Some streets' zones are missing")
                                        .Show();
                            }
                            else
                            {
                                ShowStatusMessage("Excel file modified successfully.", 5);
                                if(!this.ContainsFocus)
                                    new ToastContentBuilder()
                                        .AddText("Finished")
                                        .AddText("Your excel file is ready")
                                        .Show();
                            }   
                            
                        }
                            
                        foreach (string editedFile in modifiedFilesPath)
                        {
                            // edited file is eiter modified exc el file or text file containing missing streets
                            if (File.Exists(editedFile))
                            {
                                if (!this.ContainsFocus) // if the current app isnt this then show it to prevent freezing
                                {
                                    SetForegroundWindow(hWnd);
                                    ShowWindow(hWnd, 5); // 5 = show                     
                                }
                            
                                SaveFile(editedFile);
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        ShowStatusMessage("", 0);
                        if (ex is ArgumentNullException || ex is SQLiteException || ex is ArgumentException)
                            showError(ex.Message);
                        else if (ex is MissingFieldException)
                            showWarning(ex.Message);
                        else
                        {
                            showError(ex.Message);
                            //showError("Something went wrong the app will close now");
                            Application.Exit();
                        }
                    }
                }

                else if (!Directory.Exists(dataPath))
                    showError("Data folder is missing.");

                progressBar1.Visible = false;
                this.Cursor = Cursors.Default;
            }
        }

        public static void showError(string msg)
        {
            MessageBox.Show(msg,
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error // .Warning for Warning  
                                             //MessageBoxIcon.Error // for Error 
                                             //MessageBoxIcon.Information  // for Information
                                             //MessageBoxIcon.Question // for Question
                    );
        }

        public static void showWarning(string msg)
        {
            MessageBox.Show(msg,
                        "Warning",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning // .Warning for Warning  
                                             //MessageBoxIcon.Error // for Error 
                                             //MessageBoxIcon.Information  // for Information
                                             //MessageBoxIcon.Question // for Question
                    );
        }
        private string selectFile()
        {
            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;
                }
            }

            return filePath;
        }

        public static void SaveFile(string path)
        {
            if (path.Length == 0) return;
            
            using (SaveFileDialog Save = new SaveFileDialog())
            {
                Save.Filter = "Excel Files (*.xlsx)|*.xlsx";
                Save.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                Save.Title = "Save The Modified Excel file";
                if (path.EndsWith(".txt"))
                {
                    Save.Title = "Save The Missing Streets text file";
                    Save.Filter = "Text Files (*.txt)|*.txt";
                }
                    
                Save.FileName = path.Substring(path.LastIndexOf('\\') + 1);
                
                if (Save.ShowDialog() == DialogResult.OK)
                {
                    if (File.Exists(Save.FileName))
                        File.Delete(Save.FileName);
                    File.Move(path, Save.FileName);
                }
                else
                {
                    File.Delete(path);
                }
            }
        }

        private async void clearData_Click(object sender, EventArgs e)
        {
            if (progressBar1.Visible == false)
            {
                var confirmResult = MessageBox.Show("Are you sure you want to delete all the data?",
                                     "Confirm Delete!!",
                                     MessageBoxButtons.OKCancel);
                if (confirmResult == DialogResult.OK)
                {
                    this.Cursor = Cursors.WaitCursor;
                    progressBar1.Visible = true;
                    progressBar1.Style = ProgressBarStyle.Marquee;
                    ShowStatusMessage("Clearing data...", 0);
                    try
                    {
                        await Task.Run(() => EditExcelFile.ClearData());
                        ShowStatusMessage("Data cleared successfully.", 5);
                    }
                    catch (SQLiteException ex)
                    {
                        ShowStatusMessage("Couldn't delete", 5);
                        showError(ex.Message);
                    }
                    catch (MissingFieldException)
                    {
                        ShowStatusMessage("No data to clear", 5);
                    }
                    finally
                    {
                        progressBar1.Visible = false;
                        this.Cursor = Cursors.Default;
                    }
                }
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                if (progressBar1.Visible)
                {
                    if (MessageBox.Show(this,
                            "Are you sure you want to close the app?",
                            "Closing App",
                            MessageBoxButtons.OKCancel,
                            MessageBoxIcon.Question) == DialogResult.Cancel)
                    {
                        // cancel the form closing if necessary
                        e.Cancel = true;
                    }
                    else Application.Exit();

                }
            }
            
        }

        private void ShowStatusMessage(string message, int time)
        {
            timer.Stop();
            
            statusBar.Text = message;
            
            if (time == 0)
            {
                return;
            }

            timer.Interval = time * 1000; //time seconds
            timer.Enabled = true;
        }
    }
}
