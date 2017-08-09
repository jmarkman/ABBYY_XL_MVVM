using System;
using ABBYY_XL_MVVM.Model;
using OfficeOpenXml;
using ABBYY_XL_MVVM.Components;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace ABBYY_XL_MVVM.ViewModel
{
    /// <summary>
    /// ViewModel for the ABBYY data
    /// </summary>
    public class ABBYYDataViewModel
    {
        // Create an instance of the Model for the ViewModel to communicate with and use
        public ABBYYDataModel AppData { get; set; }
        // These are all the private variables needed to export the results to an Excel spreadsheet
        private string[] headers =
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
        private readonly string userDesktop = $@"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}";
        private readonly string dateToday = DateTime.Now.ToString("MM-dd-yyyy");

        /// <summary>
        /// Instantiate the Model for the ViewModel's usage
        /// </summary>
        public ABBYYDataViewModel()
        {
            AppData = new ABBYYDataModel();
        }

        // Access the FillABBYYDataGrid method within the Model to acquire the data and pass it to the View
        public void FillGrid()
        {
            AppData.FillABBYYDataGrid();
        }

        /// <summary>
        /// Exports the results of the search and any user edits to an Excel spreadsheet
        /// </summary>
        public void ExportABBYYDataGrid()
        {
            // Layer of abstraction so I don't have to refer to AppData.ABBYYData every time
            DataTable data = AppData.ABBYYData;
            if (data == null)
                return;

            // Create a list of rows and declare a variable to hold a singular row
            List<WorkstationRow> allRows = new List<WorkstationRow>();
            WorkstationRow currentRow;
            // Nested foreach to iterate through each row and store each cell in a WorkstationRow
            foreach (DataRow row in data.Rows)
            {
                currentRow = new WorkstationRow();
                foreach (DataColumn column in data.Columns)
                {
                    string currentCell = row[column].ToString();
                    currentRow.Add(currentCell);
                }
                allRows.Add(currentRow);
            }

            // Create excel file for results, and if that file already exists, delete it and remake it
            FileInfo excelFile = new FileInfo(userDesktop + $"Control Number [{AppData.ControlNumber}] ({dateToday})");
            if (excelFile.Exists)
            {
                excelFile.Delete();
                excelFile = new FileInfo(userDesktop + $"Control Number [{AppData.ControlNumber}] ({dateToday})");
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
                string excelBookName = $"Control Number [{AppData.ControlNumber}] ({dateToday}).xlsx";
                Byte[] sheetAsBinary = pkg.GetAsByteArray();
                File.WriteAllBytes(Path.Combine(userDesktop, excelBookName), sheetAsBinary);
            }
        }
    }
}
