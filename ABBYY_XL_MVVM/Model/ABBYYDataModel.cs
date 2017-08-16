using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using OfficeOpenXml;
using ABBYY_XL_MVVM.Components;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using System.Data.SQLite;
using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Net;
using System.Globalization;
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
        private DataGridCellInfo _cellInfo;
        // The following are all the private variables needed for the PPC lookup
        private readonly string requestUri = "https://maps.googleapis.com/maps/api/geocode/json?sensor=false&address=";
        private readonly string apiKey = Properties.Resources.APIKey;
        // The following are all the private variables needed to export the results to an Excel spreadsheet
        // WKFC workstation headers
        private readonly string[] headers =
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
        private readonly string userDesktop = $@"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}";
        // Today's date
        private readonly string dateToday = DateTime.Now.ToString("MM-dd-yyyy");


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

        public DataGridCellInfo CellInfo
        {
            get { return _cellInfo; }
            set
            {
                _cellInfo = value;
                OnPropertyChanged(nameof(CellInfo));
                // Debug to see if it works
                MessageBox.Show($"Col Index: {_cellInfo.Column.DisplayIndex.ToString()}");
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

        /// <summary>
        /// Exports the results of the search and any user edits to an Excel spreadsheet
        /// </summary>
        public void ExportABBYYDataGrid()
        {
            // Layer of abstraction so I don't have to refer to ABBYYAppData.ABBYYData every time
            if (ABBYYData == null)
                return;

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
            FileInfo excelFile = new FileInfo(userDesktop + $"Control Number [{ControlNumber}] ({dateToday})");
            if (excelFile.Exists)
            {
                excelFile.Delete();
                excelFile = new FileInfo(userDesktop + $"Control Number [{ControlNumber}] ({dateToday})");
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
                for (int i = 0; i < headers.Length; i++)
                {
                    sheet.Cells[headerRow, i + 1].Value = headers[i];
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
                string excelBookName = $"Control Number [{ControlNumber}] ({dateToday}).xlsx";
                Byte[] sheetAsBinary = pkg.GetAsByteArray();
                File.WriteAllBytes(Path.Combine(userDesktop, excelBookName), sheetAsBinary);
            }
        }

        /// <summary>
        /// Based on the City and State, acquires the county and populates the Protection Class Code box if empty
        /// </summary>
        public void PPCLookup()
        {
            /*
             * Implementation:
             * 
             * For each row in the columns City and State:
             *  Pass the city and state to google's geocoding api and retrieve the county
             *  Take the county and query the PPC SQLite database for the corresponding PPC code
             *  Take the PPC code and assign its value to the Protection Code column of the corresponding row
             */
            foreach (DataRow abbyyRow in ABBYYData.Rows)
            {
                string city = abbyyRow["City"].ToString();
                string community = (abbyyRow["County"].ToString().Equals("") ? abbyyRow["State"].ToString() : abbyyRow["County"].ToString());

                abbyyRow["Protection Code"] = GetPPC(city, community);
            }
        }

        public string GetCountyFromGeocode(string cityState)
        {
            string county = "";
            string url = $"{requestUri}{cityState}{apiKey}";
            using (WebClient webClient = new WebClient())
            {
                var json = webClient.DownloadString(url);
                JObject geocodeResults = JObject.Parse(json);
                if ((string)geocodeResults["status"] != "OK")
                    county = "";

                county = geocodeResults["results"][0]["address_components"]
                         .Where(x => (string)x["types"][0] == "administrative_area_level_2")
                         .Select(y => y["long_name"])
                         .FirstOrDefault()
                         .ToString();
            }
            return county;
        }

        public string GetPPC(string city, string community)
        {
            string ppcCode = "";
            string conn = Properties.Resources.ConnectString;
            string cmd = $@"select code from ppcCodes where (Community like '{city}' and State like '{community}') or (Community like '{city}' and County like '{community}')";
            using (SqlConnection connect = new SqlConnection(conn))
            {
                connect.Open();
                SqlCommand sqlCmd = new SqlCommand(cmd, connect);
                SqlDataReader readCodes = sqlCmd.ExecuteReader();
                if (!readCodes.HasRows)
                    ppcCode = "";
                while (readCodes.Read())
                    ppcCode = readCodes["Code"].ToString();
            }
            return ppcCode;
        }

        /// <summary>
        /// Writes out common address abbreviations in their full form.
        /// </summary>
        // TODO: Test that the shorthand method behaves the way I think it will. 
        public void ExpandShorthand()
        {
            // Establish dictionary of shorthand keys and their expanded values
            Dictionary<string, string> shorthand = new Dictionary<string, string>()
            {
                { "ave", "Avenue" }, { "ave.", "Avenue" }, { "ave,", "Avenue," },
                { "blvd", "Boulevard" }, { "blvd.", "Boulevard" }, { "blvd,", "Boulevard," },
                { "blvd.,", "Boulevard,"}, { "rd", "Road" }, { "rd.", "Road" }, { "rd,", "Road," },
                { "rd.,", "Road,"}, { "dr", "Drive" }, { "dr.", "Drive" }, { "dr,", "Drive," },
                { "dr.,", "Drive,"}, { "ln", "Lane" }, { "ln.", "Lane" }, { "ln,", "Lane," },
                { "ln.,", "Lane,"}
            };
            // Create a new instance of the TextInfo class, which is a child class of CultureInfo.
            // This class instance will allow us to use the ToTitleCase method below.
            TextInfo txt = new CultureInfo("en-US", false).TextInfo;
            // Take the address, make it lowercase for identification, split the address up by spaces, and place the pieces in a list
            foreach (DataRow item in ABBYYData.Rows)
            {
                // Take the address, make it lowercase for identification, split the address up by spaces, and place the pieces in a list
                List<string> splitAddress = item["Street 1"].ToString().ToLower().Split(' ').ToList();
                foreach (var kvp in shorthand)
                {
                    // If the shorthand from the dictionary exists in the address part list
                    if (splitAddress.Contains(kvp.Key)) 
                    {
                        // Find it and update it
                        int pieceLocation = splitAddress.IndexOf(kvp.Key);
                        splitAddress[pieceLocation] = kvp.Value;
                    }
                    item["Street 1"] = txt.ToTitleCase(string.Join(" ", splitAddress));
                }
            }
        }

        /// <summary>
        /// Expands the abbreviations for cardinal directions into their proper word form
        /// </summary>
        // TODO: Implement the cardinal direction method and test it
        public void ExpandCardinalDirection()
        {
            /*
             * Implementation:
             * 
             * For each row in the datatable
             *  For each cardinal direction
             *   If the text (converted to lowercase) in the Street 1 cell contains a full cardinal direction
             *    The method doesn't need to do anything (set needsExpansion to false)
             *    
             * If !needsExpansion
             *  return
             * Else
             *  Using regex, find the singular letter used to represent the cardinal direction
             *  If the match is successful
             *   Based on a switch statement that determines if the letter is at the end of the string or within
             *   replace the letter with the full word
             */ 
        }

        // Implementation of INotifyPropertyChanged interface
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
