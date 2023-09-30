using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace fileSeek
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            progressBar1.Value = 0;
            richTextBox1.Text = "";
            string directoryPath = textBox2.Text;
            string[] docxFiles = Directory.GetFiles(directoryPath, "*");



            var text = textBox1.Text;

            var i = 0;

            foreach (string file in docxFiles)
            {
                i++;
                if (!file.Contains("~") || !file.Contains("$"))
                {
                    

                    if (file.Contains(".xlsx"))
                    {

                        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

                        // Open the Excel file
                        Workbook workbook = excelApp.Workbooks.Open(file);

                        // Get the first worksheet in the workbook
                        Worksheet worksheet = workbook.Sheets[1];

                        // Get the used range of the worksheet
                        Range usedRange = worksheet.UsedRange;

                        // Loop through each cell in the used range
                        foreach (Range cell in usedRange.Cells)
                        {
                            // Check if the cell value contains the search keyword
                            if (cell.Value != null && cell.Value.ToString().Contains(text))
                            {
                                richTextBox1.Text = richTextBox1.Text + "\n" + file + " - " + cell.Address;

                            }
                        }

                        // Close the workbook and release resources
                        workbook.Close();
                        excelApp.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    }
                    else if (file.Contains(".docx"))
                    {            // Create an instance of the Word application
                        Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();    // Open the Word document
                        Document wordDoc = wordApp.Documents.Open(file);

                        // Read the content of the document
                        string content = wordDoc.Content.Text;

                        if (content.Contains(text))
                        {
                            //listView1.Items.Add(file);
                            richTextBox1.Text += "\n" + file;
                        }
                        wordDoc.Close();
                    }
                    else
                    {
                        richTextBox1.Text += "\nفرمت فایل پشتیبانی نمی شود " + file + "\n";
                    }

                    
                }

                progressBar1.Value += (int)Math.Round((decimal)((100 / docxFiles.Length)));

                if (i == docxFiles.Length) progressBar1.Value = 100;

                label4.Text = progressBar1.Value.ToString() + "%";

            }

            richTextBox1.Text = richTextBox1.Text + "\n" + "پایان";


        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            //folderBrowserDialog.SelectedPath = @"C:\";
            folderBrowserDialog.Description = "Select a folder";
            folderBrowserDialog.ShowNewFolderButton = true;
            folderBrowserDialog.ShowDialog();
            textBox2.Text = folderBrowserDialog.SelectedPath;


        }

        
    }



}
