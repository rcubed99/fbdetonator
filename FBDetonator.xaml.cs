using System;
using System.Collections.Generic;
using System.ComponentModel;
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

namespace fbdetonator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ExcelReader.getExcelFile();
            MessageBox.Show("File Processed!");
        }

        public void SetStatusText(string text)
        {
            this.StatusText.Text = text;
            Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Background, new Action(delegate { this.UpdateLayout(); }));
        }

        public void SetRowCount(int count)
        {
            this.RowCount.Text = count.ToString();
            Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Background, new Action(delegate { this.UpdateLayout(); }));
        }

        /*public void SetPageDonateCount(int count)
        {
            this.PageDonationCount.Text = count.ToString();
            Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Background, new Action(delegate { this.UpdateLayout(); }));
        }*/

        public void SetPostDonateCount(int count)
        {
            this.PostDonationCount.Text = count.ToString();
            Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Background, new Action(delegate { this.UpdateLayout(); }));
        }

        /*public void SetFundraiserDonateCount(int count)
        {
            this.FundraiserDonationCount.Text = count.ToString();
            Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Background, new Action(delegate { this.UpdateLayout(); }));
        }*/
    }
}
