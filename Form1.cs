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

namespace Grile_2024
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        string nextLine(StreamReader sr)
        {
            string v = "";
            while (v == "") v = sr.ReadLine();
            return v;
        }



        void readFile(string file)
        {
            StreamReader sr = new StreamReader(file);
            Excel.Application app = new Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Add();
            Excel._Worksheet ws = workbook.ActiveSheet;
            ws.Range[ws.Cells[1,2],ws.Cells[1,6]].Merge();
            ws.Range[ws.Cells[2, 2], ws.Cells[2, 6]].Merge();
            string title = file.Substring(file.LastIndexOf('\\') + 1);
            title.Remove(title.Length-3);

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

            v = nextLine(sr);
            while (!sr.EndOfStream)
            {

                ws.Cells[i, 1] = v + nextLine(sr);


                for (int j = 2; j <= 4; j++)
                {
                    v = nextLine(sr);

                    ws.Cells[i, j] = v;
                }

                v = nextLine(sr);
                if(v==null || v == "")
                {
                    continue;

                }
                ws.Cells[i, 6] = v.Substring(v.Length - 1);

                i++;

                v = nextLine(sr);

                /*
                ws.Cells[i, 1] = sr.ReadLine() +"\n"+ sr.ReadLine();
                for (int j = 2; j <= 4; j++)
                {
                    ws.Cells[i, j] = sr.ReadLine();
                }

                string rasp = sr.ReadLine();

                ws.Cells[i, 6] = rasp.Substring(rasp.Length - 1);
                i++;
                sr.ReadLine();
                */
            }

            string loc = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Grile";
            if (!Directory.Exists(loc))
            {
                Directory.CreateDirectory(loc);
            }


            workbook.SaveAs(loc+"\\"+title.Replace("txt","xlsx"));
            workbook.Close();
            sr.Close();
            app.Quit();


        }





        private void Form1_Load(object sender, EventArgs e)
        {
            string res = Environment.CurrentDirectory+"\\Grile";
            if (Directory.Exists(res))
            {
                foreach(string file in Directory.GetFiles(res))
                {
                    readFile(file);
                }
            }
            else
            {
                throw new DirectoryNotFoundException();
            }
        }
    }
}
