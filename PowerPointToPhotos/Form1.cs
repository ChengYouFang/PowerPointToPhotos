using System;
using System.ComponentModel;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Core;

namespace PowerPointToPhotos
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private static string type;
        private static string filePath;
        private static string fileName;
        private static string outPath;

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "PowerPoint 97|*.ppt|PowerPoint 2007|*.pptx";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                type = dialog.FileName.Substring(dialog.FileName.LastIndexOf('.') + 1);
                filePath = dialog.FileName;
                fileName = Path.GetFileNameWithoutExtension(filePath);
                outPath = dialog.FileName.Substring(0, dialog.FileName.LastIndexOf('\\') + 1);

                BackgroundWorker worker = new BackgroundWorker();
                worker.WorkerReportsProgress = true;
                worker.WorkerSupportsCancellation = true;
                worker.DoWork += backgroundWorker_DoWork;
                worker.RunWorkerCompleted += backgroundWorker_RunWorkerCompleted;
                worker.RunWorkerAsync();
            }
        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                MessageBox.Show("以取消");
            }
            else if (e.Error != null)
            {
                MessageBox.Show("error");
            }
            else
            {
                MessageBox.Show("輸出完成");
            }
        }

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                var pp = new Microsoft.Office.Interop.PowerPoint.Application();
                int index = 0;

                if (type.Equals("ppt"))
                {
                    var ppt = pp.Presentations.Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

                    foreach (Microsoft.Office.Interop.PowerPoint.Slide s in ppt.Slides)
                    {
                        s.Export(Path.Combine(outPath, string.Format("{0}{1}.jpg", fileName, index)), "jpg", Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
                        index++;
                    }

                    ppt.Close();
                }
                else if (type.Equals("pptx"))
                {
                    var ppt = pp.Presentations.Open2007(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

                    foreach (Microsoft.Office.Interop.PowerPoint.Slide s in ppt.Slides)
                    {
                        s.Export(Path.Combine(outPath, string.Format("{0}{1}.jpg", fileName, index)), "jpg", Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
                        index++;
                    }

                    ppt.Close();
                }
            }
            catch (NullReferenceException ex)
            {
            }
            finally
            {
                type = string.Empty;
                filePath = string.Empty;
                fileName = string.Empty;
                outPath = string.Empty;
            }
        }
    }
}