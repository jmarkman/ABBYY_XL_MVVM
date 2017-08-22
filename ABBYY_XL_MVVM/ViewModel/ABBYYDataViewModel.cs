using System.Windows;
using System.Windows.Controls;
using ABBYY_XL_MVVM.Model;
using System.Threading;

namespace ABBYY_XL_MVVM.ViewModel
{
    /// <summary>
    /// ViewModel for the ABBYY data
    /// </summary>
    public class ABBYYDataViewModel
    {
        // Create an instance of the Model for the ViewModel to communicate with and use
        public ABBYYDataModel ABBYYAppData { get; set; }

        /// <summary>
        /// Instantiate the Model for the ViewModel's usage
        /// </summary>
        public ABBYYDataViewModel()
        {
            ABBYYAppData = new ABBYYDataModel();
        }
        
        // Access the FillABBYYDataGrid method within the Model to acquire the data and pass it to the View
        public void FillABBYYDataGrid()
        {
            ABBYYAppData.FillABBYYDataGrid();
        }

        public void ExportABBYYDataGrid()
        {
            ABBYYAppData.ExportABBYYDataGrid();
            MessageBox.Show("The file has been exported to you desktop", "Export Complete");
        }

        public void PPCLookup()
        {
            Thread thread = new Thread(ABBYYAppData.PPCLookup);
            thread.Start();
        }

        // See TODO in ABBYYDataModel
        //public void DeleteRow(DataGridRowEventArgs eventArgs)
        //{
        //    ABBYYAppData.DeleteRow(eventArgs);
        //}

        public void About()
        {
            MessageBox.Show("ABBYY-XL for WKFC Underwriting Managers\n\nDeveloped by Alfred Long and Jonathan Markman\nIcons sourced from Fatcow Web Hosting (https://www.fatcow.com/free-icons)", "About ABBYY-XL");
        }
    }
}
