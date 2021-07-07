using System;
using System.Collections.Generic;
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
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace CompareTables
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string firstFile = null;
        private string secondFile = null;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void File1_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog(); //открываем меню выбора файла
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "Microsoft Excel (*.xls*)|*.xls*";
            ofd.Title = "Выберите 1-й файл для сверки";
            
            if (!(bool)ofd.ShowDialog()) //если юзер не выбрал файл
            {
                MessageBox.Show("Вы не выбрали файл для загрузки", "Загрузка данных...", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            firstFile = ofd.FileName; // - путь к файлу
            File1.Content = ofd.FileName.Substring(ofd.FileName.LastIndexOf('\\') + 1);
        }

        private void File2_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog(); //открываем меню выбора файла
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "Microsoft Excel (*.xls*)|*.xls*";
            ofd.Title = "Выберите 2-й файл для сверки";

             if (!(bool)ofd.ShowDialog()) //если юзер не выбрал файл
            {
                MessageBox.Show("Вы не выбрали файл для загрузки", "Загрузка данных...", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            secondFile = ofd.FileName; // - путь к файлу
            
            

            File2.Content = ofd.FileName.Substring(ofd.FileName.LastIndexOf('\\')+1);
      


        }


        //нажатие на кнопку когда все файлы загружены!!
        private void Action_Click(object sender, RoutedEventArgs e)
        {
            if ((firstFile == null) || (secondFile == null))
            {
                MessageBox.Show("Вы не выбрали файлы", "Загрузка данных...", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            /*
            Excel.Application xlApp1 = new Excel.Application(); //создаём приложение Excel
            Excel.Workbook xlWB1 = xlApp1.Workbooks.Open(firstFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing); //открываем наш файл           
            Excel.Worksheet xlSht1 = (Excel.Worksheet)xlWB1.Worksheets.get_Item(1); //или так xlSht = xlWB.ActiveSheet //активный лист


            //            int x = int.Parse(xlSht1.Cells[1, 1]);
            var b = xlSht1.Range["A1"].Value2;
                int d = 0;

            /*
            Excel.Application excApp = new Excel.Application(); // экземпляр приложения
            Excel.Workbook excWb; //создаём экземпляр рабочей книги
            Excel.Worksheet excWS; //экземпляр листа
            excWb = excApp.Workbooks.Add();
            excWS = (Excel.Worksheet)excWb.Worksheets.get_Item(1);

            for (int j = 1; j < 10; j++)
            {
                excWS.Cells[1, j] = j;
            }
            Excel.Range rng = excWS.Range["A2"];
            rng.Formula = "=SUM(A1:L1)";
            rng.FormulaHidden = false;
            */
            //xlApp.Visible = true;
            //xlApp.UserControl = true;

            
            Excel.Application xlApp1 = new Excel.Application(); //создаем excel-приложение
            Excel.Application xlApp2 = new Excel.Application(); //создаем excel-приложение
            Excel.Application newApp = new Excel.Application(); //создаем excel-приложение
            Excel.Workbook xlWB1; //excel-файл
            Excel.Worksheet xlWS1; //excel-лист
            xlWB1 = xlApp1.Workbooks.Open(firstFile); //инициализируем переменные нашими фйлами
            xlWS1 =(Excel.Worksheet) xlWB1.Worksheets["Лист1"]; //или как там? кароче, индекс - название листа

            Excel.Workbook xlWB2; //excel-файл
            Excel.Worksheet xlWS2; //excel-лист
            xlWB2 = xlApp2.Workbooks.Open(secondFile); //инициализируем переменные нашими фйлами
            xlWS2 = (Excel.Worksheet)xlWB2.Worksheets["Лист1"]; //или как там? кароче, индекс - название листа

            Excel.Workbook newWb = newApp.Workbooks.Add();
            Excel.Worksheet newWs = (Excel.Worksheet)newWb.Worksheets.get_Item(1);
            string ind;
            for (int i=1; i<10; i++)
            {

                ind = "A" + i.ToString();
                if ((int)xlWS1.Range[ind].Value2 >= (int)xlWS2.Range[ind].Value2)
                {
                    
                    newWs.Cells[i,1] = 1; //свойство Cells только для записи, а Range только для чтения 
                    
                } else
                {
                    newWs.Cells[i, 1] = 0;
                    (newWs.Cells[i, 1] as Excel.Range).Interior.ColorIndex = 37;
                    // newWs.get_Range(i, 1).Font.Color = Excel.XlRgbColor.rgbMediumVioletRed;
                }

            }

            newApp.Visible = true;
            newApp.UserControl = true;
            

        }

        //начало описания драгндропа
        private void field1Enter(object sender, DragEventArgs e) 
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                Field1.Fill = Brushes.White;
            }
            //File1.Visibility = Visibility.Collapsed;
        }

        private void field2Enter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                Field2.Fill = Brushes.White;
            }
            
           // File2.Visibility = Visibility.Hidden;
        }

        private void field1Leave(object sender, DragEventArgs e)
        {
            Field1.Fill = Brushes.Silver;
            //File1.Visibility = Visibility.Visible;
        }
        private void field2Leave(object sender, DragEventArgs e)
        {
            Field2.Fill = Brushes.Silver;
            //File2.Visibility = Visibility.Visible;
        }

        private void File2_Drop(object sender, DragEventArgs e)
        {
            var rr = e.Data.GetFormats();
            var same = e.Data.GetData("FileNameW");
            string name = ((string[])same)[0];
            string ext = null;
            if (name.IndexOf('.') != -1)
            {
                ext = name.Substring(name.LastIndexOf('.') + 1, name.Length - name.LastIndexOf('.') - 1); //получение расширения                                                                                                //
            }
            else
            ext = "folder";
            if ((ext == "xls") || (ext == "xlsx"))
            {
                secondFile = name;
                File2.Content = name.Substring(name.LastIndexOf('\\') + 1); //имя файла
            } else
            {

                MessageBox.Show("Загруженный файл имеет некорректное расширение!\nПоддерживаемые расширения: xls, xlsx.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
          
            }

            Field2.Fill = Brushes.Silver;


        }

        private void File1_Drop(object sender, DragEventArgs e)
        {
            var rr = e.Data.GetFormats();
            var same = e.Data.GetData("FileNameW");
            string name = ((string[])same)[0];
            string ext = null;
            if (name.IndexOf('.') != -1)
            {
                ext = name.Substring(name.LastIndexOf('.') + 1, name.Length - name.LastIndexOf('.') - 1); //получение расширения                                                                                                //
            }
            else
                ext = "folder";
            if ((ext == "xls") || (ext == "xlsx"))
            {
                firstFile = name;
                File1.Content = name.Substring(name.LastIndexOf('\\') + 1); //имя файла
            }
            else
            {

                MessageBox.Show("Загруженный файл имеет некорректное расширение!\nПоддерживаемые расширения: xls, xlsx.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }

            Field1.Fill = Brushes.Silver;


        }

    }
}
