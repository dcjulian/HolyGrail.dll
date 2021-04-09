
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
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

namespace DataReadExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            String filePath = @"C:\Users\john\Documents\netflix-shows\netflix_shows.xlsx";
            try {

                IWorkbook workbook = null;
                FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                if (filePath.IndexOf(".xlsx") > 0)
                    workbook = new XSSFWorkbook(fs);
                else if (filePath.IndexOf(".xls") > 0)
                    workbook = new HSSFWorkbook(fs);


                Console.WriteLine(filePath.IndexOf(".xlsx"));
                Console.WriteLine(filePath.IndexOf(".xls"));
                
                ISheet sheet = workbook.GetSheetAt(0);

                if(sheet != null)
                {
                    int rowCount = sheet.LastRowNum;

                 for(int i=1; i<=rowCount; i++)
                    {
                        IRow curROw = sheet.GetRow(i);


                    }
                }

            }
            catch(Exception exception)
            {
                Console.WriteLine(exception.Message);
            }
        }
    }
}
