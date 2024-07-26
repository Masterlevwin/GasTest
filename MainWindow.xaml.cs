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
        private readonly DefaultDialogService dialogService = new DefaultDialogService();

        public ObservableCollection<TubeObject> TubeObjects = new ObservableCollection<TubeObject>();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void OpenFile(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dialogService.OpenFileDialog() == true && dialogService.FilePaths != null)
                {
                    OpenFile(dialogService.FilePaths);
                }
            }
            catch (Exception ex) { dialogService.ShowMessage(ex.Message); }
        }

        private void OpenFile(string[] paths)
        {
            TubeObjects.Clear();

            for (int i = 0; i < paths.Length; i++)
            {
                if (Path.GetExtension(paths[i]) == ".csv") ImportObjectsCSV(paths[i]);
                else if (Path.GetExtension(paths[i]) == ".xls" || Path.GetExtension(paths[i]) == ".xlsx") ImportObjectsXLS(paths[i]);
                else MessageBox.Show($"Формат файла {Path.GetExtension(paths[i])} не поддерживается");
            }

            ShowTubeObjectsGrid();
        }

        private void ImportObjectsXLS(string path)
        {
            using (FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                {
                    DataSet result = reader.AsDataSet();
                    DataTable table = result.Tables[0];

                    for (int i = 1; i < table.Rows.Count; i++)
                    {
                        if ($"{table.Rows[i].ItemArray[0]}" == "") break;

                        TubeObject tube = new TubeObject()
                        {
                            Name = $"{table.Rows[i].ItemArray[0]}",
                            Distance = Parser($"{table.Rows[i].ItemArray[1]}"),
                            Angle = Parser($"{table.Rows[i].ItemArray[2]}"),
                            Width = Parser($"{table.Rows[i].ItemArray[3]}"),
                            Height = Parser($"{table.Rows[i].ItemArray[4]}")
                        };

                        if ($"{table.Rows[i].ItemArray[5]}".Contains("yes")) tube.IsDefect = true;

                        TubeObjects.Add(tube);
                    }
                }
            }
        }

        private void ImportObjectsCSV(string path)
        {
            using (TextFieldParser parser = new TextFieldParser(path))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(";");
                while (!parser.EndOfData)
                {
                    string[] cells = parser.ReadFields();

                    if (cells[0].Contains("Name") || cells[0] == "") continue;

                    TubeObject tube = new TubeObject()
                    {
                        Name = $"{cells[0]}",
                        Distance = Parser($"{cells[1]}"),
                        Angle = Parser($"{cells[2]}"),
                        Width = Parser($"{cells[3]}"),
                        Height = Parser($"{cells[4]}")
                    };

                    if ($"{cells[5]}".Contains("yes")) tube.IsDefect = true;

                    TubeObjects.Add(tube);
                }
            }
        }

        private void ShowTubeObjectsGrid()
        {
            TubeObjectsGrid.ItemsSource = TubeObjects;

            TubeObjectsGrid.Columns[0].Header = "Название";
            TubeObjectsGrid.Columns[1].Header = "Гор";
            TubeObjectsGrid.Columns[2].Header = "Верт";
            TubeObjectsGrid.Columns[3].Header = "Шир";
            TubeObjectsGrid.Columns[4].Header = "Выс";
            TubeObjectsGrid.Columns[5].Header = "Дефект";
        }

        public static float Parser(string data)     //обёртка для парсинга float-значений
        {
            //если число одновременно содержит и запятую, и точку - удаляем запятую
            if (data.Contains(',') && data.Contains('.')) data = data.Replace(",", "");

            //если есть запятая в строке, заменяем ее на точку, и парсим строку в число
            if (float.TryParse(data.Replace(',', '.'), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands,
                System.Globalization.CultureInfo.InvariantCulture, out float f)) return f;
            else return 0;
        }

        private void Exit(object sender, RoutedEventArgs e)
        {

        }

        private void ShowGrafics(object sender, SelectionChangedEventArgs e)
        {
            Grafics();
        }

        private void Grafics()
        {
            if (TubeObjectsGrid.SelectedItem is TubeObject tube)
            {
                RectGrid.Children.Clear();

                Polygon rect = new Polygon();
                rect.Points = new PointCollection
                {               
                    new Point(tube.Distance * 10, tube.Angle * 10),
                    new Point((tube.Distance + tube.Width) * 10, tube.Angle * 10),
                    new Point((tube.Distance + tube.Width) * 10, (tube.Angle + tube.Height) * 10),
                    new Point(tube.Distance * 10, (tube.Angle + tube.Height) * 10)
                };

                rect.Stroke = Brushes.Red;
                rect.StrokeThickness = 2;

                RectGrid.Children.Add(rect);
            }
            else return;
        }
    }
}
