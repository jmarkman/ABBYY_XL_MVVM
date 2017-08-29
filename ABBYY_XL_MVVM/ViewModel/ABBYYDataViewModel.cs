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
            /*
             * Forseen bug:
             * If the user, for some reason, makes a new search as the method is active,
             * the method keeps running because of the nature of threads. What results is the
             * population of the new search's PPC column with the PPC lookup for the previous
             * search that was currently running. If the search left off at row 4, the lookup
             * will start inserting at row 5 when the new results return. This is probably a
             * great place for a critical error if the new search returns fewer rows than the 
             * previous (i.e., the previous search had 10 rows and the new one has 3, and the
             * previous PPC lookup left off at row 5)
             */
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
