using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using ExcelReader;
using ExcelReader.Models;
using Microsoft.Win32;
using Newtonsoft.Json;

namespace ExcelToJsonConverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            _excelFileMgr = new ExcelFileManager();

            DataContext = this;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            _path = string.Empty;
            filePath.Content = _path;

            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                _path = openFileDialog.FileName;
                filePath.Content = _path;

                availableSheets.Clear();
                foreach (SheetInfo sheet in _excelFileMgr.GetFileSheets(_path))
                {
                    availableSheets.Add(sheet);
                }
            }
        }

        public ObservableCollection<SheetInfo> availableSheets { get; set; } = new ObservableCollection<SheetInfo>();

        private List<SheetInfo> sheetsToGetDataFrom = new List<SheetInfo>();
        private ExcelFileManager _excelFileMgr;
        private string _path;

        private void listView1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            foreach (SheetInfo item in e.RemovedItems)
            {
                sheetsToGetDataFrom.Remove(item);
            }

            foreach (SheetInfo item in e.AddedItems)
            {
                sheetsToGetDataFrom.Add(item);
            }
        }

        private void ConvertToJSON_Click(object sender, RoutedEventArgs e)
        {
            Dictionary<string, List<Dictionary<string, string>>> content = _excelFileMgr.GetPageContent(sheetsToGetDataFrom);

            string json = JsonConvert.SerializeObject(content);

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                FileName = $"export-{DateTime.UtcNow.ToString("yyyyMMdd")}",
                DefaultExt = ".json",
                Filter = "Json files (*.json)|*.json|Text files (*.txt)|*.txt"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                File.WriteAllText(saveFileDialog.FileName, json);
            }
        }
    }
}
