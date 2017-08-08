using ABBYY_XL_MVVM.Model;

namespace ABBYY_XL_MVVM.ViewModel
{
    /// <summary>
    /// ViewModel for the ABBYY data
    /// </summary>
    public class ABBYYDataViewModel
    {
        // Create an instance of the Model for the ViewModel to communicate with and use
        public ABBYYDataModel AppData { get; set; }

        public ABBYYDataViewModel()
        {
            AppData = new ABBYYDataModel();
        }

        // Access the FillGrid method within the Model to acquire the data and pass it to the View
        public void FillGrid()
        {
            AppData.FillGrid();
        }
    }
}
