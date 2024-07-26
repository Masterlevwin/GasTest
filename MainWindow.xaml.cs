using ExcelDataReader;
using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using Path = System.IO.Path;

namespace GasTest
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly DefaultDialogService dialogService = new DefaultDialogService();               //сервис диалоговых окон

        public ObservableCollection<TubeObject> TubeObjects = new ObservableCollection<TubeObject>();   //коллекция объектов задания

        public MainWindow()
        {
            InitializeComponent();
        }

        private void OpenFiles(object sender, RoutedEventArgs e)     //обработчик кнопки открытия файла / файлов
        {
            try
            {
                if (dialogService.OpenFileDialog() == true && dialogService.FilePaths != null)
                {
                    OpenFiles(dialogService.FilePaths);
                }
            }
            catch (Exception ex) { dialogService.ShowMessage(ex.Message); }
        }

        private void OpenFiles(string[] paths)                      //открытие файла / файлов
        {
            TubeObjects.Clear();

            for (int i = 0; i < paths.Length; i++)
            {
                //взависимости от расширения файла запускаем соответствующий парсер
                if (Path.GetExtension(paths[i]) == ".csv") ImportObjectsCSV(paths[i]);
                else if (Path.GetExtension(paths[i]) == ".xls" || Path.GetExtension(paths[i]) == ".xlsx") ImportObjectsXLS(paths[i]);
                else MessageBox.Show($"Формат файла {Path.GetExtension(paths[i])} не поддерживается");
            }

            ShowTubeObjectsGrid();
        }

        private void ImportObjectsXLS(string path)                  //парсер таблиц Excel в форматах xls / xlsx
        {
            using (FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                //здесь используем готовую библиотеку ExcelDataReader
                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                {
                    DataSet result = reader.AsDataSet();
                    DataTable table = result.Tables[0];

                    //перебираем строки полученного DataTable
                    for (int i = 1; i < table.Rows.Count; i++)
                    {
                        if ($"{table.Rows[i].ItemArray[0]}" == "") break;

                        //считываем значения ячеек в свойства объекта задания
                        TubeObject tube = new TubeObject()
                        {
                            Name = $"{table.Rows[i].ItemArray[0]}",                 //название
                            Distance = Parser($"{table.Rows[i].ItemArray[1]}"),     //горизонтальная координата
                            Angle = Parser($"{table.Rows[i].ItemArray[2]}"),        //вертикальная координата
                            Width = Parser($"{table.Rows[i].ItemArray[3]}"),        //горизонтальный размер объекта
                            Height = Parser($"{table.Rows[i].ItemArray[4]}")        //вертикальный размер объекта
                        };

                        if ($"{table.Rows[i].ItemArray[5]}".Contains("yes")) tube.IsDefect = true;  //дефектный ли объект

                        TubeObjects.Add(tube);      //добавляем созданный объект класса в коллекцию
                    }
                }
            }
        }

        private void ImportObjectsCSV(string path)                  //парсер содержимого файла csv
        {
            //здесь используем стандартную библиотеку Microsoft.VisualBasic
            using (TextFieldParser parser = new TextFieldParser(path))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(";");
                while (!parser.EndOfData)
                {
                    string[] cells = parser.ReadFields();

                    if (cells[0].Contains("Name") || cells[0] == "") continue;      //пропускаем первую и пустые строки

                    TubeObject tube = new TubeObject()
                    {
                        Name = $"{cells[0]}",                 //название
                        Distance = Parser($"{cells[1]}"),     //горизонтальная координата
                        Angle = Parser($"{cells[2]}"),        //вертикальная координата
                        Width = Parser($"{cells[3]}"),        //горизонтальный размер объекта
                        Height = Parser($"{cells[4]}")        //вертикальный размер объекта
                    };

                    if ($"{cells[5]}".Contains("yes")) tube.IsDefect = true;  //дефектный ли объект

                    TubeObjects.Add(tube);      //добавляем созданный объект класса в коллекцию
                }
            }
        }

        private void ShowTubeObjectsGrid()                          //метод отображения коллекции объектов в таблице DataGrid
        {
            TubeObjectsGrid.ItemsSource = TubeObjects;

            TubeObjectsGrid.Columns[0].Header = "Название";
            TubeObjectsGrid.Columns[1].Header = "Гор";
            TubeObjectsGrid.Columns[2].Header = "Верт";
            TubeObjectsGrid.Columns[3].Header = "Шир";
            TubeObjectsGrid.Columns[4].Header = "Выс";
            TubeObjectsGrid.Columns[5].Header = "Дефект";
        }

        public static float Parser(string data)                     //обёртка для парсинга float-значений
        {
            //если число одновременно содержит и запятую, и точку - удаляем запятую
            if (data.Contains(',') && data.Contains('.')) data = data.Replace(",", "");

            //если есть запятая в строке, заменяем ее на точку, и парсим строку в число
            if (float.TryParse(data.Replace(',', '.'), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands,
                System.Globalization.CultureInfo.InvariantCulture, out float f)) return f;
            else return 0;
        }

        private void ShowGrafics(object sender, SelectionChangedEventArgs e)    //обработчик события выбора строки в DataGrid
        {
            Grafics();
        }

        private void Grafics()                                      //метод, в котором строится графическое изображение объекта задания
        {
            if (TubeObjectsGrid.SelectedItem is TubeObject tube)
            {
                RectGrid.Children.Clear();                          //очищаем предыдущее изображение

                Polygon rect = new Polygon();
                rect.Points = new PointCollection                   //создаем линию по точкам, полученным из свойств объекта задания
                {               
                    new Point(tube.Distance * 10, tube.Angle * 10),
                    new Point((tube.Distance + tube.Width) * 10, tube.Angle * 10),
                    new Point((tube.Distance + tube.Width) * 10, (tube.Angle + tube.Height) * 10),
                    new Point(tube.Distance * 10, (tube.Angle + tube.Height) * 10)
                };

                rect.Stroke = Brushes.Red;                          //окрашиваем линию в красный
                rect.StrokeThickness = 2;                           //делаем линию толще

                RectGrid.Children.Add(rect);
            }
            else return;
        }


        private void Exit(object sender, RoutedEventArgs e)         //выход из программы
        {
            Environment.Exit(0);
        }
    }
}
