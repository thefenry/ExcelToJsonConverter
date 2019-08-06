using System.Collections.Generic;
using System.Windows;
using ExcelReader;
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

            string path = @"";
            ExcelFileManager em = new ExcelFileManager(path);
            Dictionary<string, int> sheets = em.GetFileSheets();

            var content = em.GetPageContent(sheets);

            var json = JsonConvert.SerializeObject(content);
        }
    }
}
