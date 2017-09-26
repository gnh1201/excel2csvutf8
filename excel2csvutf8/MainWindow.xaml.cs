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
using Microsoft.Win32;
using System.Windows.Threading;

namespace excel2csvutf8
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            createComboboxItems();
        }

        private void btnBrowse1_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "MS Excel Files(*.xls, *.xlsx)|*.xls;*.xlsx";
            openFile.DefaultExt = "xlsx";
            openFile.ShowDialog();

            if (openFile.FileNames.Length > 0)
            {
                foreach (string filename in openFile.FileNames)
                {
                    this.txtOldpath.Text = filename;
                }
            }
        }

        private void btnBrowse2_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Filter = "CSV File(*.csv)|*.csv";
            saveFile.DefaultExt = "csv";
            saveFile.ShowDialog();

            if (saveFile.FileNames.Length > 0)
            {
                foreach (string filename in saveFile.FileNames)
                {
                    this.txtNewpath.Text = filename;
                }
            }
        }

        private void createComboboxItems()
        {
            String[] regionsKeys = { "auto", "dos", "kr", "kr949", "jp", "cn" };
            String[] regions = { "Auto", "Western European(DOS)", "Asian-KR", "Asian-KR949", "Asian-JP", "Asian-CN" };
            for (int i = 0; i < regions.Length; i++)
            {
                ComboboxItem item = new ComboboxItem();
                item.Text = regions[i];
                item.Value = regionsKeys[i];
                this.chRegion.Items.Add(item);
            }

            this.chRegion.SelectedIndex = 0;
        }

        private void btnExec_Click(object sender, RoutedEventArgs e)
        {
            if (this.txtNewpath.Text == "" || this.txtOldpath.Text == "")
            {
                MessageBox.Show("Oops! Please check your form.");
            } else
            {
                string message = "Do you want run this task?";
                string caption = "Confirm task";

                MessageBoxButton buttons = MessageBoxButton.YesNo;
                MessageBoxResult result;
                result = MessageBox.Show(message, caption, buttons);

                if (result == MessageBoxResult.Yes)
                {
                    this.DataRun();
                }
                else
                {
                    MessageBox.Show("Task is cancelled.");
                }
            }
        }

        private void execDatacAsync()
        {
            Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate
            {
                DataProcessing dataproc = new DataProcessing();
                String filePath = this.txtOldpath.Text;
                String saveFilePath = this.txtNewpath.Text;

                ComboBox cmb = (ComboBox)this.chRegion;
                int selectedIndex = cmb.SelectedIndex;
                ComboboxItem selectedObject = (ComboboxItem)cmb.SelectedValue;
                String selectedRegion = (String)selectedObject.Value;

                dataproc.parse(filePath, saveFilePath, selectedRegion);
            }));
        }

        private async void DataRun()
        {

            var task1 = Task<int>.Run(() => execDatacAsync());
            await task1;

            MessageBox.Show("Well done.");
        }
    }
}
