using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace CreateTrainingFilesWindows
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string json = "{\"classifiers\":[";

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            if (excelApp != null)
            {
                Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(@"C:\Excel\Data.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Sheets[1];

                Microsoft.Office.Interop.Excel.Range excelRange = excelWorksheet.UsedRange;
                int rowCount = excelRange.Rows.Count;
                int colCount = excelRange.Columns.Count;

                List<string> list = new List<string>();
                for (int i = 2; i <= rowCount; i++)
                {
                    Microsoft.Office.Interop.Excel.Range label = (excelWorksheet.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range);
                    string stringLabel = label.Value.ToString();
                    if (list.Contains(stringLabel) == false)
                    {
                        list.Add(stringLabel);
                    }
                }
                for(int i = 0; i < list.Count; i++)
                {
                    json += "{\"name\":\"" + list[i] + "\"},";
                }
                //remove the , from last item and close file
                json = json.Remove(json.Length - 1);
                json += "],\"documents\":[";

                //i = 2 to start so we skip headers
                for (int i = 2; i <= rowCount; i++)
                {
                    Microsoft.Office.Interop.Excel.Range content = (excelWorksheet.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range);
                    Microsoft.Office.Interop.Excel.Range label = (excelWorksheet.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range);
                    string stringContent = content.Value.ToString();
                    string stringLabel = label.Value.ToString();
                    //label1.Text = label1.Text + " " + cellValue;
                    string fileName = "file" + (i - 1).ToString() + ".txt";
                    createTxtFile(fileName, stringContent);
                    json += "{\"location\":\"" + fileName + "\",\"language\":\"en-us\",\"classifiers\":[{\"classifierName\":\"" + stringLabel + "\"}]},";
                }
                //remove the , from last item and close file
                json = json.Remove(json.Length - 1) + "]}";
                createTxtFile("model.json", json);

                label1.Text = "Training files created. Now upload them to Azure Storage.";

                excelWorkbook.Close();
                excelApp.Quit();
            }
        }

        void createTxtFile(string name, string content)
        {
            string fileName = @"C:\Excel\Training\" + name;

            try
            {
                // Check if file already exists. If yes, delete it.     
                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                }

                // Create a new file     
                using (FileStream fs = File.Create(fileName))
                {
                    // Add some text to file    
                    Byte[] fileContent = new UTF8Encoding(true).GetBytes(content);
                    fs.Write(fileContent, 0, fileContent.Length);
                }
            }
            catch (Exception Ex)
            {
                throw Ex;
            }
        }
    }
}
