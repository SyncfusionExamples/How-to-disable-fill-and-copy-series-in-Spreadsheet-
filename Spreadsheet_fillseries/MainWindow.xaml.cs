﻿using Syncfusion.UI.Xaml.Spreadsheet;
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

namespace Spreadsheet_fillseries
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            spreadsheet.Create(3);
            spreadsheet.WorkbookLoaded += Spreadsheet_WorkbookLoaded;
            spreadsheet.WorksheetAdded += Spreadsheet_WorksheetAdded;
        }

        private void Spreadsheet_WorksheetAdded(object sender, Syncfusion.UI.Xaml.Spreadsheet.Helpers.WorksheetAddedEventArgs args)
        {
            spreadsheet.ActiveGrid.FillSeriesController.AllowFillSeries = false;
        }

        private void Spreadsheet_WorkbookLoaded(object sender, Syncfusion.UI.Xaml.Spreadsheet.Helpers.WorkbookLoadedEventArgs args)
        {
            foreach(var sheet in args.GridCollection)
            {
                sheet.FillSeriesController.AllowFillSeries = false;             
            }
        }
    }
}
