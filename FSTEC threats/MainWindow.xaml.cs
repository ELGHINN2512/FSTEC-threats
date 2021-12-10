using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
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

using ExcelDataReader;
using Microsoft.Win32;
using System.Net;
using System.ComponentModel;

namespace FSTEC_threats
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ThreatTable threatTable = new ThreatTable("thrlist.xlsx");

        public MainWindow()
        {
            InitializeComponent();
        }

        private void MenuItemFull_Click(object sender, RoutedEventArgs e)
        {
            if(threatTable == null)
            {
                threatTable = new ThreatTable("thrlist.xlsx");
            }
            else
            {
                DbGrig.ItemsSource = threatTable.СreateDataViewWithFullInformation();
            }
        }

        private void MenuItemAbbreviated_Click(object sender, RoutedEventArgs e)
        {
            if (threatTable == null)
            {
                DbGrig.ItemsSource = threatTable.СreateTableWithAbbreviatedInformation();
            }
            else
            {
                DbGrig.ItemsSource = threatTable.СreateTableWithAbbreviatedInformation();
            }
        }

        private void MenuItemUpdated_Click(object sender, RoutedEventArgs e)
        {
            threatTable.Update();
        }

    }
}
