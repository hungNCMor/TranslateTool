using Microsoft.Win32;
using System;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using TranslateLib.Interface;

namespace TranslateJpToVi
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ITranslateExcel _translate;
        ITranslateFile _translateFile;
        public MainWindow(ITranslateExcel translate, ITranslateFile translateFile)
        {
            _translate = translate;
            InitializeComponent();
            _translateFile = translateFile;
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            var path = PathTxt.Text.Trim();
            if (path.Length == 0 || !System.IO.File.Exists(path))
            {
                MessageBox.Show("Nhập đường dẫn hợp lệ hoặc chọn file");
            }
            else
            {
                Overlay.Visibility = Visibility.Visible;
                await Task.Run(async () =>
                  {
                      var fileName = path.Split('.').Last();
                      switch (fileName)
                      {
                          case "xlsx":
                          case "csv":
                          case "xls":
                           await   _translate.TranslateExcelByPathSavePath(path);
                              break;
                          case "pptx":
                          case "ppt":
                           await _translateFile.TranslateFileByPathSavePath(path);
                              break;
                      }
                      //...rest of code
                  });

                Overlay.Visibility = Visibility.Hidden;
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            var res = dlg.ShowDialog();
            if (res == true)
            {
                PathTxt.Text = dlg.FileName;
            }
        }


        private void ComboBox_SelectionChanged_2(object sender, SelectionChangedEventArgs e)
        {

        }
    }
}
