using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using OfficeOpenXml;
using ABBYY_XL_MVVM.Components;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using System;
using System.Linq;
using System.Net;
using System.Windows.Controls;
using System.Windows; // MessageBox for debug purposes, remove when ready to go forward


namespace ABBYY_XL_MVVM.Model
{
    public class ABBYYDataModel : INotifyPropertyChanged
    {
        private string _controlNumber; // Control Number of the submission
        private DataTable _abbyyData; // Data associated with the submission
        // I have this to try to get the cell row/col index on click for PPC lookup on a singular location
        // Possible that I will remove this entirely if it doesn't work out
        // private DataGridCellInfo _cellInfo;
        // The following are all the private variables needed for the PPC lookup
        private readonly string _requestUri = "https://maps.googleapis.com/maps/api/geocode/json?sensor=false&address=";
        private readonly string _apiKey = Properties.Resources.APIKey;
        // The following are all the private variables needed to export the results to an Excel spreadsheet
        // WKFC workstation headers
        private readonly string[] _headers =
        {
            "Loc #", "Bldg #", "Physical Building #", "Single Physical Building #",
            "Street 1", "Street 2", "City", "State", "Zip", "County", "Building Value",
            "Business Personal Property", "Busines Income", "Misc Real Property",
            "TIV", "# Units", "Building Description", "WKFC Major Occupancy", "WKFC Detailed Occupancy", "LRO",
            "ClassCodeDesc", "Building Usage", "Construction Type", "Dist. To Fire Hydrant (Feet)",
            "Dist. To Fire Station (Miles)", "Prot Class", "# Stories", "# Basements",
            "Year Built", "Sq Ftg", "Wiring Year", "Plumbing Year", "Roofing Year",
            "Heating Year", "Fire Alarm Type", "Burglar Alarm Type", "Sprinkler Alarm Type",
            "Sprinkler Wet/Dry", "Sprinkler Extent"
        };
        // User desktop
        private readonly string _userDesktop = $@"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}";
        // Today's date
        private readonly string _dateToday = DateTime.Now.ToString("MM-dd-yyyy");


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

        // Comment out until ready to actually utilize this
        //public DataGridCellInfo CellInfo
        //{
        //    get { return _cellInfo; }
        //    set
        //    {
        //        _cellInfo = value;
        //        OnPropertyChanged(nameof(CellInfo));
        //        // Debug to see if it works
        //        MessageBox.Show($"Col Index: {_cellInfo.Column.DisplayIndex.ToString()}");
        //    }
        //}

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

                ConstructionTypeMatchesWorkstation();
            }
        }

        /// <summary>
        /// Exports the results of the search and any user edits to an Excel spreadsheet
        /// </summary>
        public void ExportABBYYDataGrid()
        {
            if (ABBYYData == null)
                return;

            // TODO: I might not even need the custom list implementation, but it can't hurt to have it?

            // Create a list of rows and declare a variable to hold a singular row
            List<WorkstationRow> allRows = new List<WorkstationRow>();
            WorkstationRow currentRow;
            // Nested foreach to iterate through each row and store each cell in a WorkstationRow
            foreach (DataRow row in ABBYYData.Rows)
            {
                currentRow = new WorkstationRow();
                foreach (DataColumn column in ABBYYData.Columns)
                {
                    string currentCell = row[column].ToString();
                    currentRow.Add(currentCell);
                }
                allRows.Add(currentRow);
            }

            // Create excel file for results, and if that file already exists, delete it and remake it
            FileInfo excelFile = new FileInfo(_userDesktop + $"Control Number [{ControlNumber}] ({_dateToday})");
            if (excelFile.Exists)
            {
                excelFile.Delete();
                excelFile = new FileInfo(_userDesktop + $"Control Number [{ControlNumber}] ({_dateToday})");
            }

            using (ExcelPackage pkg = new ExcelPackage())
            {
                // The header row is all of the WKFC headers for the workstation
                // dataRowStart denotes the row right after the header
                // Most likely something that can be changed/refactored?
                int headerRow = 1;
                int dataRowStart = 2;

                // Make a new worksheet in the excel workbook and make sure we're using that sheet
                pkg.Workbook.Worksheets.Add("ABBYY Results");
                ExcelWorksheet sheet = pkg.Workbook.Worksheets[1];
                sheet.Name = "ABBYY Results";

                // Write the headers to the excel sheet
                for (int headerPosition = 0; headerPosition < _headers.Length; headerPosition++)
                {
                    sheet.Cells[headerRow, headerPosition + 1].Value = _headers[headerPosition];
                }

                // Write the actual data gleaned from the DataGrid and save it 
                foreach (var row in allRows)
                {
                    for (int rowIndex = 0; rowIndex < row.RowLength; rowIndex++)
                    {
                        sheet.Cells[dataRowStart, rowIndex + 1].Value = row.DataAtIndex(rowIndex);
                    }
                    dataRowStart++;
                }
                string excelBookName = $"Control Number [{ControlNumber}] ({_dateToday}).xlsx";
                Byte[] sheetAsBinary = pkg.GetAsByteArray();
                File.WriteAllBytes(Path.Combine(_userDesktop, excelBookName), sheetAsBinary);
            }
        }

        /// <summary>
        /// Based on the City and State, acquires the county and populates the Protection Class Code box if empty
        /// </summary>
        public void PPCLookup()
        {
            foreach (DataRow abbyyRow in ABBYYData.Rows)
            {
                // Assign the city and county/state to appropriate variables
                string city = abbyyRow["City"].ToString();
                string community = (abbyyRow["County"].ToString().Equals("") ? abbyyRow["State"].ToString() : abbyyRow["County"].ToString());

                // Set that row's Protection Code column to the results of the method, GetPPC()
                if (string.IsNullOrEmpty(abbyyRow["Protection Code"].ToString()))
                {
                    abbyyRow["Protection Code"] = GetPPC(city, community);
                }
            }
        }

        /// <summary>
        /// Based on the city-state combo, return the county of that city-state combo to the user
        /// </summary>
        /// <param name="cityState">The city and state of a given location as a string</param>
        /// <returns>The county as a string</returns>
        public string GetCountyFromGeocode(string cityState)
        {
            // Build storage variable and Geocode URL
            string county = "";
            string url = $"{_requestUri}{cityState}{_apiKey}";
            // Based on the json returned, LINQ through it to find the appropriate county
            using (WebClient webClient = new WebClient())
            {
                var json = webClient.DownloadString(url);
                JObject geocodeResults = JObject.Parse(json);
                if ((string)geocodeResults["status"] != "OK")
                    county = "";

                /*
                 * Google's Geocode results are unnecessarily difficult to serialize to .NET objects because
                 * 1. Not every piece of JSON returned has the same structure
                 * 2. Copying and pasting the JSON as a class in VS creates a class with lots of little subclasses and arrays of those little subclasses
                 * 
                 * All we really need is to get the county from the results. Relevant results will always appear in the "results" json as the first item in the results array, and each relevant result will have an array of "address components"
                 */ 
                county = geocodeResults["results"][0]["address_components"]
                         .Where(x => (string)x["types"][0] == "administrative_area_level_2")
                         .Select(y => y["long_name"])
                         .FirstOrDefault()
                         .ToString();
            }
            return county;
        }

        /// <summary>
        /// Connects to the database to fetch the PPC code based on the provided city and community (state or county)
        /// </summary>
        /// <param name="city">The city value from the datatable</param>
        /// <param name="community">The county or state from the datatable</param>
        /// <returns>The PPC as a string</returns>
        public string GetPPC(string city, string community)
        {
            string ppcCode = ""; // Variable to store the PPC code
            string conn = Properties.Resources.ConnectString;
            // Query for accessing the codes
            string cmd = 
                $@"select Code from PPCData where (Community like '{city}' and State like '{community}') or (Community like '{city}' and County like '{community}')";
            using (SqlConnection connect = new SqlConnection(conn))
            {
                connect.Open();
                SqlCommand sqlCmd = new SqlCommand(cmd, connect);
                SqlDataReader readCodes = sqlCmd.ExecuteReader();
                // If there is no code available, skip this by setting the return value to a blank string
                if (!readCodes.HasRows)
                    ppcCode = "";
                while (readCodes.Read())
                    // Set the variable equal to the value returned from the database 
                    ppcCode = readCodes["code"].ToString();
            }
            return ppcCode;
        }

        /// <summary>
        /// Based on the construction type, append a number to the start of the construction type
        /// string so that it matches the expected input for the workstation
        /// </summary>
        public void ConstructionTypeMatchesWorkstation()
        {
            // Assign numbers to construction types as they are on the workstation
            Dictionary<string, string> workstationConstrType = new Dictionary<string, string>()
            {
                { "Frame", "1." },
                { "Joisted Masonry", "2." },
                { "Non-Combustible", "3." },
                { "Masonry, Non-Combustible", "4." },
                { "Modified Fire Resistive", "5." },
                { "Fire Resistive", "6." }
            };

            // Iterate through the DataTable and perform the changes as necessary
            foreach (DataRow abbyyRow in ABBYYData.Rows)
            {
                // Abstract the value in the column so that we don't have to type out what's being assigned
                // to this variable every time we want to access that value
                string abbyyConstrTypeValue = abbyyRow["Construction Type"].ToString();

                foreach (KeyValuePair<string, string> kvp in workstationConstrType)
                {
                    if (abbyyConstrTypeValue.Equals(kvp.Key))
                    {
                        abbyyRow["Construction Type"] = $"{kvp.Value} {abbyyConstrTypeValue}";
                    }
                    else if (abbyyConstrTypeValue.Equals("None Provided"))
                    {
                        abbyyRow["Construction Type"] = "";
                    }
                }
            }
        }

        // TODO: Implement row deletion option or row select to let user hit "delete" key to delete
        //public void DeleteRow(DataGridRowEventArgs eventArgs)
        //{
        //    ABBYYData.Rows.RemoveAt(eventArgs.Row.GetIndex());
        //}

        // Implementation of INotifyPropertyChanged interface
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
