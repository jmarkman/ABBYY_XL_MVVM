using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ABBYY_XL_MVVM.Model;

namespace ABBYYXL_UnitTest
{
    [TestClass]
    public class ModelUnitTest
    {
        // No error should be thrown if the control number is null
        [TestMethod]
        public void TestFillABBYYDataGrid_ReturnIfNull()
        {
            ABBYYDataModel model = new ABBYYDataModel()
            {
                ControlNumber = null
            };
            model.FillABBYYDataGrid();
        }

        // No error shoul be thrown if the control number is an empty string
        [TestMethod]
        public void TestFillABBYYDataGrid_ReturnIfEmptyString()
        {
            ABBYYDataModel model = new ABBYYDataModel()
            {
                ControlNumber = ""
            };
            model.FillABBYYDataGrid();
        }

        // Prove that the grid is being filled
        // We should be getting specific values back for this entry
        [TestMethod]
        public void TestFillABBYYDataGrid()
        {
            ABBYYDataModel model = new ABBYYDataModel()
            {
                ControlNumber = "1191167"
            };
            model.FillABBYYDataGrid();
            var rows = model.ABBYYData.Rows;
            Assert.AreEqual("Farm to Market Road 482", rows[0][4]);
        }

        // Check that we're getting the county after all is said and done
        [TestMethod]
        public void TestGetCountyFromGeocode()
        {
            ABBYYDataModel model = new ABBYYDataModel();

            string testCounty = model.GetCountyFromGeocode("Melville, NY");
            Assert.AreEqual("Suffolk County", testCounty);
        }

        // Check that the method GetPPC is fetching the PPC from the database
        [TestMethod]
        public void TestGetPPC()
        {
            ABBYYDataModel model = new ABBYYDataModel();
            string ppc = model.GetPPC("Booth", "AL");
            Assert.AreEqual("5", ppc);
        }

        [TestMethod]
        public void TestIfGetPPCSkipsCellsWithContent()
        {
            /*
             * The control number 1193093 has almost all of its locations in PA, 
             * with the county being Berks. The PPC for Berks County is 4.
             * If we manually assign the PPC for row 1 as 5, the method should 
             * skip over it since that cell has content and assign the appropriate
             * value in the PPC column for the following row.
             */
            ABBYYDataModel model = new ABBYYDataModel
            {
                ControlNumber = "1193093"
            };
            model.FillABBYYDataGrid();
            var rows = model.ABBYYData.Rows;
            rows[0]["Protection Code"] = "5";
            model.PPCLookup();
            Assert.AreEqual("5", rows[0]["Protection Code"].ToString());
            Assert.AreEqual("4", rows[1]["Protection Code"].ToString());
        }
    }
}
