using Microsoft.Win32;
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
using TranslateLib;
using TranslateLib.Interface;

namespace TranslateJpToVi
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ITranslateExcel _translate;
        public MainWindow(ITranslateExcel translate)
        {
            _translate = translate;
            InitializeComponent();
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
              await  Task.Run(() =>
                {
                    _translate.TranslateExcelByPathSavePath(path);
                    //...rest of code
                });
              
                //await Task.Delay(10000 );
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
    }
}
