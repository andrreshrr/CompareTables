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

        private string GetColumnName(int col_number) //преобразует номер в аналог обозначения столбца екселя (1-А, 2-B, 3-C, ... , 28 - AB) 
        {
            string result;
            if (col_number > 0)
            {
                int alphabets = (col_number - 1) / 26;
                int remainder = (col_number - 1) % 26;
                result = ((char)('A' + remainder)).ToString();
                if (alphabets > 0)
                {
                    result = GetColumnName(alphabets) + result;
                }
            }
            else result = null;

            return result;
        }

        private int GetRows(Excel.Worksheet xlWS1)
        {
            //КОД ДЛЯ НАХОЖДЕНИЯ ЧИСЛА СТРОК И СТОЛБЦОВ, его по-хорошему в функцию запихать какую
            int rows = 1;
            Excel.Range current_elem = xlWS1.Range["A" + rows.ToString()];
       
            while (xlWS1.Range["A" + rows.ToString()].Value2 != null)
            {
                rows++;

            }
            return rows;
            
            //rows и cols - число строк+1 и число столбцов+1, для удобства в форе
        }

        private int GetCols(Excel.Worksheet xlWS1)
        {
            Excel.Range current_elem = xlWS1.Range["A1"];
            int cols = 1;
            while (current_elem.Value2 != null)
            {
                cols++;
                current_elem = current_elem.Next;

            }
            return cols;
        }

        //нажатие на кнопку когда все файлы загружены!!
        private void Action_Click(object sender, RoutedEventArgs e)
        {
            if ((firstFile == null) || (secondFile == null))
            {
                MessageBox.Show("Вы не выбрали файлы", "Загрузка данных...", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

      
            Excel.Application newApp = new Excel.Application(); //создаем excel-приложение
            Excel.Workbook xlWB1; //excel-файл
            Excel.Worksheet xlWS1; //excel-лист
            xlWB1 = newApp.Workbooks.Open(firstFile); //инициализируем переменные нашими фйлами
            xlWS1 =(Excel.Worksheet) xlWB1.Worksheets.get_Item(1); //в скобочках индекс листа, индексация начинается с единицы

            Excel.Workbook xlWB2; //excel-файл
            Excel.Worksheet xlWS2; //excel-лист
            xlWB2 = newApp.Workbooks.Open(secondFile); //инициализируем переменные нашими фйлами
            xlWS2 = (Excel.Worksheet)xlWB2.Worksheets.get_Item(1); //в скобочках индекс листа, индексация начинается с единицы

           

            int cols1 = GetCols(xlWS1);     //cols1 и rows1 число солбцов+1 и строк+1 первого листа первой таблицы
            int rows1 = GetRows(xlWS1);
            int cols2 = GetCols(xlWS2);
            int rows2 = GetRows(xlWS2);

            int nErow = 1; //строка в новой таблице

            string ind1, ind2;
            
            int i = 0, j=0;

          

            var curCell1 = xlWS1.Range["A1"];
            List<int> data1 = new List<int>();
            while (curCell1.Value2 != null)
            {
                i++;
                ind1 = "A" + i.ToString();
                data1.Add((int)curCell1.Value2);
                curCell1 = xlWS1.Range[ind1];
                
            } 

            var curCell2 = xlWS2.Range["A1"];
            List<int> data2 = new List<int>();
            while (curCell2.Value2 != null)
            {
                j++;
                ind2 = "A" + j.ToString();
                data2.Add((int)curCell2.Value2);
                curCell2 = xlWS2.Range[ind2];

            }

            var res1 = data2.Except(data1);
            var res2 = data1.Except(data2);
            //УЖЕ ЭТИ ДВА РЕЗА ОТПРАВЛЯТЬ В ИТОГОВУЮ ТАБЛИЦУ
            Excel.Workbook newWb = newApp.Workbooks.Add();
            Excel.Worksheet newWs1 = (Excel.Worksheet)newWb.Worksheets.get_Item(1);
           

            i = 1;
            foreach (var item in res1)
            {
                if (i < rows1)
                {
                    newWs1.Cells[i, 1] = item;
                    i++;
                }
                else break;

            }

            j = 1;
            
            newWb.Worksheets.Add();
            Excel.Worksheet newWs2 = (Excel.Worksheet)newWb.Worksheets.get_Item(1);

            foreach (var item in res2)
            {
                if (j < rows2)
                {
                    newWs2.Cells[j, 1] = item;
                    j++;
                }
                else break;

            }

            /*
            Excel.Workbook newWb = newApp.Workbooks.Add();
            Excel.Worksheet newWs = (Excel.Worksheet)newWb.Worksheets.get_Item(1);
            string ind;
            for (int j = 1; j < cols; j++)
            {
                for (int i = 1; i < rows; i++)
                {

                    ind = GetColumnName(j) + i.ToString();
                    if ((int)xlWS1.Range[ind].Value2 >= (int)xlWS2.Range[ind].Value2)
                    {
                        newWs.Cells[i, j] = 1; //свойство Cells только для записи, а Range только для чтения 

                        (newWs.Cells[i,j] as Excel.Range).Interior.ColorIndex = 4; //изменили цвет 

                    }
                    else
                    {
                        newWs.Cells[i, j] = 0;
                        (newWs.Cells[i, j] as Excel.Range).Interior.ColorIndex = 3; //red
                    }

                }

            } */

            xlWB1.Close(false); //false значит не сохранять изменения, хотя мы ничего и не изменяли, но пусть
            xlWB2.Close(false);
           

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
