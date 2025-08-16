using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace Grile_2024
{
    public partial class Form1 : Form
    {
        private string errorString = "�";


        private string savePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Grile\\Grila2donev2.xlsx";


        public Form1()
        {
            InitializeComponent();
        }

   


        void readFile(string file)
        {
            StreamReader sr = new StreamReader(file);
            Excel.Application app = new Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Add();
            Excel._Worksheet ws = workbook.ActiveSheet;
            ws.Range[ws.Cells[1,2],ws.Cells[1,6]].Merge();
            ws.Range[ws.Cells[2, 2], ws.Cells[2, 6]].Merge();
            string title = Path.GetFileNameWithoutExtension(file);

            ws.Cells[1, 1] = "Title";
            ws.Cells[1, 2] = title;
            ws.Cells[2, 2] = title;
            ws.Cells[2, 1] = "Description";
            ws.Cells[3, 1] = "Duration";
            ws.Cells[4, 1] = "Question";
            ws.Cells[3, 2] = "60";

            ws.Cells[4, 2] = "Option1";
            ws.Cells[4, 3] = "Option2";
            ws.Cells[4, 4] = "Option3";
            ws.Cells[4, 5] = "Explanation";
            ws.Cells[4, 6] = "Answer";


            int i = 5;
            string v;

            while (!sr.EndOfStream)
            {
                
                ws.Cells[i, 1] = sr.ReadLine();
                for (int j = 2; j <= 4; j++)
                {
                    ws.Cells[i, j] = sr.ReadLine().Substring(3);
                }

                //string rasp = sr.ReadLine();

                //ws.Cells[i, 6] = rasp.Substring(rasp.Length - 1);
                i++;
                
            }

            string loc = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Grile";
            if (!Directory.Exists(loc))
            {
                Directory.CreateDirectory(loc);
            }
            savePath = loc + "\\" + title + ".xlsx";
            if (File.Exists(savePath))
                File.Delete(savePath);
            

            workbook.SaveAs(savePath);
            workbook.Close(true);
            sr.Close();
            app.Quit();
            Process.Start(savePath);


        }





        private void Form1_Load(object sender, EventArgs e)
        {
          
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string res = Environment.CurrentDirectory + "\\Grile";
            if (Directory.Exists(res))
            {
                foreach (string file in Directory.GetFiles(res, "*.txt"))
                {
                    readFile(file);
                }
            }
            else
            {
                Directory.CreateDirectory(res);
                MessageBox.Show("No files found in the Grile directory. Please add some text files to process.");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FileInfo file = new FileInfo(Environment.CurrentDirectory + "\\Grile\\Grila Principal1.txt");
      
            StreamReader sr = new StreamReader(file.FullName);
            StringBuilder sb = new StringBuilder(sr.ReadToEnd());
            int i = 0;
            while(i< sb.Length)
            {
                if (sb[i] == '�')
                {
                    sb.Remove(i, 1);
                    sb.Insert(i, "ti");
                }
                if (sb[i] == '.')
                {
                    for(int j = i-1;j>=i-5 && j>=0;j--)
                    {
                        if (sb[j] == ' ')
                        {
                            sb.Remove(j, 1);
                            sb.Insert(j, "\r\n");
                            i++;
                            break;
                        }
                    }
                }

                i++;
            }
            textBox1.Text = sb.ToString();  
            sr.Close();
            File.WriteAllText(file.DirectoryName+"\\redone.txt", sb.ToString());


        }

        private void button3_Click(object sender, EventArgs e)
        {
            StreamReader sr = new StreamReader(Environment.CurrentDirectory + "\\raspPart1.txt");
            textBox1.Clear();

            int row = 5;


            Excel.Application app = new Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Open(savePath);
            Excel._Worksheet ws = workbook.ActiveSheet;






            while (!sr.EndOfStream)
            {
                string line = sr.ReadLine();
                foreach(char c in line)
                {
                    if (c >= 'a' && c <= 'z')
                    {
                        textBox1.AppendText($"{row}: {c}\r\n");
                        ws.Cells[row++, 6] = c.ToString();
                    }
                }
            }

            workbook.Save();
            workbook.Close(true);
            sr.Close();
            app.Quit();


        }

        private void button4_Click(object sender, EventArgs e)
        {
            StreamReader sr = new StreamReader(Environment.CurrentDirectory + "\\regexForta.txt");
            int row = 465;



            Excel.Application app = new Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Open(savePath);
            Excel._Worksheet ws = workbook.ActiveSheet;






            while (!sr.EndOfStream)
            {
                string line = sr.ReadLine();
                if (row == 673) row++;
                if (line.Length > 0)
                {
                    ws.Cells[row++, 6] = line;
                }

            }
            workbook.Save();
            workbook.Close(true);
            sr.Close();
            app.Quit();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Excel.Application app = new Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Open(savePath);
            Excel._Worksheet ws = workbook.ActiveSheet;
            int row = 500;
            textBox1.Clear();
            while(row <= 1034)
            {
                string text = ws.Cells[row, 6].Value2;
                if (text.Contains("\r\n"))
                {
                    textBox1.AppendText(text);
                    text = text.Replace("\r\n", "");
                    ws.Cells[row, 6].Value2 = text+"\r\n";
                }
                row++;
            }



            workbook.Save();
            workbook.Close(true);
            app.Quit();
        }
    }
}
