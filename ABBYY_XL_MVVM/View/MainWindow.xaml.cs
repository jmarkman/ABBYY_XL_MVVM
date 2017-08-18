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
using ABBYY_XL_MVVM.ViewModel;

namespace ABBYY_XL_MVVM
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly ABBYYDataViewModel _viewModel = new ABBYYDataViewModel();

        public MainWindow()
        {
            InitializeComponent();
            // Assign the data context of the window to the current ViewModel
            DataContext = _viewModel;
        }

        /// <summary>
        /// Event for the user pressing enter within the Control Number input box, calls FillABBYYDataGrid()
        /// </summary>
        private void CtrlNumInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                _viewModel.FillABBYYDataGrid();
            }
        }

        /// <summary>
        /// Toggles readonly data grid
        /// </summary>
        private void Edit_Click(object sender, RoutedEventArgs e)
        {
            if (SearchResults.IsReadOnly == true)
            {
                SearchResults.IsReadOnly = false;
                ProtCode.IsEnabled = true;
            }
            else
            {
                SearchResults.IsReadOnly = true;
                ProtCode.IsEnabled = false;
            }
        }

        /// <summary>
        /// Initiates the database search
        /// </summary>
        private void Search_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.FillABBYYDataGrid();
        }

        /// <summary>
        /// Shows "About" box. Will probably convert to "Help" later.
        /// </summary>
        private void About_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.About();
        }

        /// <summary>
        /// Exports the current datagrid as an Excel document
        /// </summary>
        private void Export_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.ExportABBYYDataGrid();
        }

        /// <summary>
        /// Fills in all Protection Class rows based on results from the database
        /// </summary>
        private void PPC_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.PPCLookup();
        }

        // See TODO in ABBYYDataModel
        //private void DeleteRow_Click(object sender, RoutedEventArgs e)
        //{
        //    //_viewModel.DeleteRow();
        //}
    }
}
