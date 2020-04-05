using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Window = System.Windows.Window;

namespace LearnEnglishCards
{
    public partial class MainWindow : Window
    {
        string[,] arr;
        int strok = 0;
        string answer = "";
        string question = "";
        bool isAnswer = false;
        string path = "English.xlsx";
        YandexTranslator yt;
        public MainWindow()
        {
            InitializeComponent();
            yt = new YandexTranslator();
            if (Environment.GetCommandLineArgs().Length > 1)
            {
                //MessageBox.Show(Environment.GetCommandLineArgs()[1].ToString());
                path = Environment.GetCommandLineArgs()[1].ToString();
            }
            
            foreach (UIElement c in LayoutRoot.Children)
            {
                if (c is System.Windows.Controls.Button)
                {
                    ((System.Windows.Controls.Button)c).Click += Button_Click;
                }
            }
            //ReadExcel();
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            string s = (string)((System.Windows.Controls.Button)e.OriginalSource).Content;
            if (s == "Translate") Translate();
            if (s == "Show") FindAndShow();

            if (s=="Next") ShowMeCards();

            if (s == "Load")
            {
                try
                {
                    Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(Directory.GetCurrentDirectory() + @"\"+""+ path + "", 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                    ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
                    var lastCell = ObjWorkSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
                    //MessageBox.Show(lastCell.Column.ToString());//Столбцы
                    //MessageBox.Show(lastCell.Row.ToString());//Строки
                    //strok = lastCell.Row;
                    //arr = new string[lastCell.Column, lastCell.Row];
                    //arr = new string[2, lastCell.Row];
                    int iLastRow = ObjWorkSheet.Cells[ObjWorkSheet.Rows.Count, "A"].End[Microsoft.Office.Interop.Excel.XlDirection.xlUp].Row;  //последняя заполненная строка в столбце А
                    //MessageBox.Show(iLastRow.ToString());
                    arr = new string[2, iLastRow];
                    strok = iLastRow;
                    var arrData = (object[,])ObjWorkSheet.Range["A1:B" + iLastRow].Value;
                    //MessageBox.Show(arrData.Length.ToString());
                    string a1 = "";
                    string a2 = "";
                    for (int d = 1; d < iLastRow+1; d++)
                    {
                        for (int i = 1; i < 3; i++)
                        {
                            //MessageBox.Show(arrData.GetValue(d, i).ToString());
                            // MessageBox.Show(arrData.GetValue(d, i).ToString().Length.ToString());
                            arr[i-1, d-1] = arrData.GetValue(d, i).ToString();
                            if (i == 1) a1 = arrData.GetValue(d, i).ToString();
                            else a2 = arrData.GetValue(d, i).ToString();

                        }
                        //MessageBox.Show(a1 + " " + a2);
                    }
                    ObjExcel.Quit();
                    messageBlock.Text = "Loading complete! Number of loading items: "+((arr.Length/2)).ToString();
                    ShowMeCards();
                }
                catch (Exception ex)
                {
                    ObjExcel.Quit();
                    MessageBox.Show(ex.Message.ToString());
                }
                finally
                {
                    ObjExcel.Quit();
                }
            }
        }
        public void Translate()
        {
            textBlock.Text = yt.Translate(textBlock.Text, "en-ru");
        }
        public void ShowMeCards()
        {
            //arr[a,b] - a (столбец), b (строка)

            Random rnd = new Random();
            int value = rnd.Next(0, strok);
            //MessageBox.Show(strok.ToString());
            question = arr[0, value].ToString();
            answer = arr[1, value].ToString();
            textBlock.Text= question;
            isAnswer = false;
        }
        public void FindAndShow()
        {
            //MessageBox.Show(arr.GetLength(1).ToString());
            if (!isAnswer)
            {
                textBlock.Text = answer;
                isAnswer = true;
            }
            else
            {
                textBlock.Text = question;
                isAnswer = false;
            }
        }
    }
}
