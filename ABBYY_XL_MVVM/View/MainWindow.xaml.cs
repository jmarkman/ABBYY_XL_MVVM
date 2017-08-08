using System.Windows;
using System.Windows.Input;
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
            DataContext = _viewModel;
        }

        /// <summary>
        /// Event for the user pressing enter within the Control Number input box, calls FillGrid()
        /// </summary>
        private void CtrlNumInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                _viewModel.FillGrid();
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
            }
            else
            {
                SearchResults.IsReadOnly = true;
            }
        }
    }
}
