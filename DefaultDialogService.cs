using Microsoft.Win32;
using System.Windows;

namespace GasTest
{
    public interface IDialogService
    {
        void ShowMessage(string message);   // показ сообщения
        string[] FilePaths { get; set; }   // путь к выбранному файлу
        bool OpenFileDialog();  // открытие файла
        bool SaveFileDialog();  // сохранение файла
    }

    public class DefaultDialogService : IDialogService
    {
        public string[] FilePaths { get; set; }

        public bool OpenFileDialog()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                Filter = "Excel-File (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                FilePaths = openFileDialog.FileNames;
                return true;
            }
            return false;
        }

        public bool SaveFileDialog()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog()
            {
                Filter = "Excel-File (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            };
            if (saveFileDialog.ShowDialog() == true)
            {
                FilePaths = saveFileDialog.FileNames;
                return true;
            }
            return false;
        }

        public void ShowMessage(string message) { MessageBox.Show(message); }
    }

}
