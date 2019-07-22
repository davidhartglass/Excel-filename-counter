using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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
using Excel = Microsoft.Office.Interop.Excel;       //Microsoft Excel 14 object in references-> COM tab

namespace Davids_Project
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 
    public partial class MainWindow : Window
    {
        private string myFile;
        Dictionary<string, string> myDictionary = new Dictionary<string, string>();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Filename_Extraction_Counter_Loaded(object sender, RoutedEventArgs e)
        {
            goButton.Visibility = Visibility.Collapsed;
            goButton.IsEnabled = true;
            myProgressBar.Visibility = Visibility.Collapsed;
        }

        private void BrowseFileExplorerBtn_Click(object sender, RoutedEventArgs e)
        {
            string extension;
            string filePath;
            string temp;

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.InitialDirectory = "C:\\Users\\DGH\\Desktop";

            openFileDialog.ShowDialog();
            temp = System.IO.Path.GetExtension(openFileDialog.FileName);

            if (openFileDialog.FileName == null || openFileDialog.FileName == "")
            {
                FilePathTextBox.Text = "No file selected... Please select a file";
            }
            else if (temp == ".xlsx" || temp == ".xls")
            {
                myFile = openFileDialog.FileName;
                FilePathTextBox.Text = myFile;

                goButton.Visibility = Visibility.Visible;
            }
        }

        private async void goButton_Click(object sender, RoutedEventArgs e)
        {
            myProgressBar.Visibility = Visibility.Visible;
            await Task.Run(() => getExcelData());
        }

        private async void getExcelData()
        {
            Excel_Integration myExcelIntegration = new Excel_Integration();

            this.Dispatcher.Invoke(() =>
            {
                goButton.IsEnabled = false;
            });

            await Task.Run(() => myDictionary = myExcelIntegration.getExcelData(myFile));

            this.Dispatcher.Invoke(() =>
            {
                goButton.IsEnabled = true;
            });
        }
    }
}