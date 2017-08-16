using ABBYY_XL_MVVM.Model;

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
        }

        public void PPCLookup()
        {
            ABBYYAppData.PPCLookup();
        }
    }
}
