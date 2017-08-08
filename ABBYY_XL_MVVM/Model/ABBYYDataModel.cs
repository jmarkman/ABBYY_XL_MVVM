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
        /// Getter/Setter for Control Number, fires OnPropertyChanged for MVVM data binding
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
        /// Getter/Setter for ABBYY data, fires OnPropertyChanged for MVVM data binding
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

        public void FillGrid()
        {
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
