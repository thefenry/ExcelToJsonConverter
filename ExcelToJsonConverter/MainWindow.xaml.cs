using System.Windows;
using Microsoft.Win32;

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

            //string path = @"C:\Users\ivans\Downloads\ESI Catalog_20190805.xlsx";
            //ExcelFileManager em = new ExcelFileManager(path);
            //Dictionary<string, int> sheets = em.GetFileSheets();

            //var content = em.GetPageContent(sheets);

            //var json = JsonConvert.SerializeObject(content);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string path = string.Empty;
            filePath.Content = path;

            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                path = openFileDialog.FileName;
                filePath.Content = path;
            }
        }
    }
}
