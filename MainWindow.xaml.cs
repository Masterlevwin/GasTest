using ExcelDataReader;
using System;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows;

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
            for (int i = 0; i < paths.Length; i++)
            {
                using (FileStream stream = File.Open(paths[i], FileMode.Open, FileAccess.Read))
                {
                    using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        DataSet result = reader.AsDataSet();
                        DataTable table = result.Tables[0];
                    }
                }
            }
        }

        private void ImportObjects(DataTable table)
        {
            TubeObjects.Clear();

            for (int i = 0; i < table.Rows.Count; i++)
            {
                if (table.Rows[i].ItemArray[0] is null || $"{table.Rows[i].ItemArray[0]}" == "") break;

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

        public static float Parser(string data)                             //обёртка для парсинга float-значений
        {
            //если число одновременно содержит и запятую, и точку - удаляем запятую
            if (data.Contains(',') && data.Contains('.')) data = data.Replace(",", "");

            //если есть запятая в строке, заменяем ее на точку, и парсим строку в число
            if (float.TryParse(data.Replace(',', '.'), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands,
                System.Globalization.CultureInfo.InvariantCulture, out float f)) return f;
            else return 0;
        }

    }
}
