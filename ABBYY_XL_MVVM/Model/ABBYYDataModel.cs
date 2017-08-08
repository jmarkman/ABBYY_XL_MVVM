using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;

namespace ABBYY_XL_MVVM.Model
{
    public class ABBYYDataModel : INotifyPropertyChanged
    {
        private string _controlNumber; // Control Number of the submission
        private DataTable _abbyyData; // Data associated with the submission

        /// <summary>
        /// The control number of the submission to search for
        /// </summary>
        public string ControlNumber
        {
            get { return _controlNumber; }
            set
            {
                _controlNumber = value;
                OnPropertyChanged(nameof(ControlNumber));
            }
        }

        /// <summary>
        /// The data associated with the provided control number
        /// </summary>
        public DataTable ABBYYData
        {
            get { return _abbyyData; }
            set
            {
                _abbyyData = value;
                OnPropertyChanged(nameof(ABBYYData));
            }
        }

        /// <summary>
        /// Queries the WKFC ABBYY database and uses SqlDataAdapter to fill ABBYYData (a DataTable object)
        /// </summary>
        public void FillABBYYDataGrid()
        {
            if (ControlNumber == null || ControlNumber == "")
                return;

            // Stored ABBYY SQL connection string in settings file
            string conn = Properties.Resources.ConnectString;
            // Calls stored procedure
            string cmd = $"exec Get140Data {ControlNumber}";
            using (SqlConnection connect = new SqlConnection(conn))
            {
                SqlCommand sqlCmd = new SqlCommand(cmd, connect);
                // SqlDataAdapter allows for database interaction with the DataTable type
                SqlDataAdapter sda = new SqlDataAdapter(sqlCmd);
                DataTable resultsDT = new DataTable("Location Information");
                sda.Fill(resultsDT); // Fill the datatable with the associated info from the database

                // Set the datatable equal to the Model's variable
                ABBYYData = resultsDT;
            }
        }

        // Implementation of INotifyPropertyChanged interface
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
