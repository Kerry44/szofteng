using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {

        Excel.Application xlApp; 
        Excel.Workbook xlWB;    
        Excel.Worksheet xlSheet;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CreateExcel();
        }

        private void CreateExcel()
        {
            try
            {
                
                xlApp = new Excel.Application();

      
                xlWB = xlApp.Workbooks.Add(Missing.Value);

             
                xlSheet = xlWB.ActiveSheet;

             
                CreateTable(); 
                
                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception ex) 
            {
                string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errMsg, "Error");

               
                xlWB.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
                xlWB = null;
                xlApp = null;
            }
        }

        private void CreateTable()
        {
            void CreateTable()
            {
                string[] fejl�cek = new string[] {
                    "K�rd�s", 
                    "1. v�lasz", 
                    "2. v�laszl", 
                    "3. v�lasz", 
                    "Helyes v�lasz", 
                    "k�p"};

                xlSheet.Cells[1, 1] = fejl�cek[0];
                Models.HajosContext hajosContext = new Models.HajosContext();
                var mindenK�rd�s = hajosContext.Questions.ToList();


            }

        }


    }
}