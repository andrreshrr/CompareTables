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
    public class DataPart: IEquatable<DataPart>
    {
        public long Id { get; set; }
        public string Name { get; set; }
        public string Number { get; set; }
        public string Total { get; set; }

        public bool Equals(DataPart other)
        {
            if (other is null)
                return false;

            return this.Id == other.Id;
        }

        //public override bool Equals(object obj) => Equals(obj as DataPart);
        public override int GetHashCode() => (Id).GetHashCode();
    }
    
    public partial class MainWindow : Window
    {
        private string firstFile = null;
        private string secondFile = null;
        public MainWindow()
        {
            InitializeComponent();

            wait1.Visibility = Visibility.Hidden;
            wait2.Visibility = Visibility.Hidden;

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
        
        
       
        
        private List <DataPart> CreateDataListFromColumn(ref Excel.Worksheet current_xlWS, string startIdCell, string startNameCell, string startNumberCell, string startTotalCell)
        {
            List<DataPart> res = new List<DataPart>();
            int i = Convert.ToInt32(startIdCell[1]) - 48; //"1"=49 => "1" - 48 = 1
            var curId = current_xlWS.Range[startIdCell].Value2;
            var curName =current_xlWS.Range[startNameCell].Value2;
            var curNumber = current_xlWS.Range[startNumberCell].Value2;
            var curTotal =current_xlWS.Range[startTotalCell].Value2;
            long inp;
            while (curId != null)
            {
                string[] ttl = Convert.ToString(curId).Split(new char[] {'-'});
                if (ttl[0] == "N")
                {
                    res.Add(new DataPart() {Id=Convert.ToInt64(ttl[1]), Name = Convert.ToString(curName), Number = Convert.ToString(curNumber) ,Total = Convert.ToString(curTotal)}); 
                }
                else if (Int64.TryParse(ttl[0],out inp))
                {
                    res.Add(new DataPart() {Id=inp,  Name = Convert.ToString(curName), Number = Convert.ToString(curNumber) ,Total = Convert.ToString(curTotal)});
                }
                i++;
                curId = current_xlWS.Range[startIdCell[0] + i.ToString()].Value2;
                curName = current_xlWS.Range[startNameCell[0] + i.ToString()].Value2;
                curNumber = current_xlWS.Range[startNumberCell[0] + i.ToString()].Value2;
                curTotal = current_xlWS.Range[startTotalCell[0] + i.ToString()].Value2;
            }
            return res;
        }
        
        //нажатие на кнопку когда все файлы загружены!!
        private void Action_Click(object sender, RoutedEventArgs e)
        {
            
            wait1.Visibility = Visibility.Visible; // делаю видимыми ректангел и лейбл А ОН НЕТ            
            wait2.Visibility = Visibility.Visible;
            
            
                if ((firstFile == null) || (secondFile == null))
            {
                MessageBox.Show("Вы не выбрали файлы", "Загрузка данных...", MessageBoxButton.OK, MessageBoxImage.Error);

                wait1.Visibility = Visibility.Hidden;
                wait2.Visibility = Visibility.Hidden;
                return;
            }

            if (MessageBox.Show($"Начать сравнение файлов: {File2.Content} и {File2.Content}?", "Вы уверены?", MessageBoxButton.YesNo)==MessageBoxResult.No){
                wait1.Visibility = Visibility.Hidden;
                wait2.Visibility = Visibility.Hidden;
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





            string startIdCell1 = text1.Text!="" ? text1.Text : "A1";
            string startIdCell2 = text2.Text!="" ? text2.Text : "A1";
            string startNameCell1 = text1name.Text!="" ? text1name.Text : "B1";
            string startNameCell2 = text2name.Text!="" ? text2name.Text : "B1";
            string startNumberCell1 = text1count.Text!="" ? text1count.Text : "C1";
            string startNumberCell2 = text2count.Text!="" ? text2count.Text : "C1";
            string startTotalCell1 = text1sum.Text!="" ? text1sum.Text : "D1";
            string startTotalCell2 = text2sum.Text!="" ? text2sum.Text : "D1";
            var data1 = CreateDataListFromColumn(ref xlWS1, startIdCell1, startNameCell1,startNumberCell1,startTotalCell1); //создаем из первой колонки лист
            var data2 = CreateDataListFromColumn(ref xlWS2, startIdCell2, startNameCell2,startNumberCell2,startTotalCell2);

            
            Excel.Workbook newWb = newApp.Workbooks.Add();
            Excel.Worksheet newWs1 = (Excel.Worksheet)newWb.Worksheets.get_Item(1);
            //string gk = "Строки " + File1.Content + ", отс. в " + File2.Content;

            
            int i =3;
            newWs1.Cells[1, 1] = "Элементы из " + File2.Content + ", которых нет в " + File1.Content;
            newWs1.Cells[2, 1] = "Индефикационные номера" ;
            newWs1.Cells[2, 2] = "Наименование" ;
            newWs1.Cells[2, 3] = "Сумма" ;
            newWs1.Cells[2, 4] = "Количество" ;
            foreach (var item in data2.Except(data1))
            {
                newWs1.Cells[i, 1] = item.Id;
                newWs1.Cells[i, 2] = item.Name;
                newWs1.Cells[i, 3] = item.Number;
                newWs1.Cells[i, 4] = item.Total;
                i++;
            }
        

            i = 3;
            
            newWb.Worksheets.Add();
            Excel.Worksheet newWs2 = (Excel.Worksheet)newWb.Worksheets.get_Item(1);
            newWs2.Cells[1, 1] = "Элементы из " + File1.Content + ", которых нет в " + File2.Content;
            newWs2.Cells[2, 1] = "Индефикационные номера" ;
            newWs2.Cells[2, 2] = "Наименование" ;
            newWs2.Cells[2, 3] = "Сумма" ;
            newWs2.Cells[2, 4] = "Количество" ;
            foreach (var item in data1.Except(data2))
            {
                newWs2.Cells[i, 1] = item.Id;
                newWs2.Cells[i, 2] = item.Name;
                newWs2.Cells[i, 3] = item.Number;
                newWs2.Cells[i, 4] = item.Total;
                i++;
            }

            /*foreach (Panel item in fullGrid.Children)
                item.Visibility = Visibility.Visible;
            warning.Visibility = Visibility.Hidden;*/
            //mainWindow.Visibility=Visibility.Visible;
            xlWB1.Close(false); //false значит не сохранять изменения, хотя мы ничего и не изменяли, но пусть
            xlWB2.Close(false);

            wait1.Visibility = Visibility.Hidden; //делаю невидимыми
            wait2.Visibility = Visibility.Hidden;
           
            newApp.Visible = true; // даём юзеру итоговый файл
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
                lineFix.Stroke = Brushes.White;
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
            lineFix.Stroke = Brushes.Silver;
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
            lineFix.Stroke = Brushes.Silver;
            Int32 c = 100;

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
