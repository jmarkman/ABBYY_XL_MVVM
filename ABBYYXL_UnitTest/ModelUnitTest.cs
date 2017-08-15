using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ABBYY_XL_MVVM.Model;

namespace ABBYYXL_UnitTest
{
    [TestClass]
    public class ModelUnitTest
    {
        [TestMethod]
        public void TestFillABBYYDataGrid_ReturnIfNull()
        {
            ABBYYDataModel testModel = new ABBYYDataModel()
            {
                ControlNumber = null
            };
            testModel.FillABBYYDataGrid();
        }

        [TestMethod]
        public void TestFillABBYYDataGrid_ReturnIfEmptyString()
        {
            ABBYYDataModel testModel = new ABBYYDataModel()
            {
                ControlNumber = ""
            };
            testModel.FillABBYYDataGrid();
        }

        [TestMethod]
        public void TestFillABBYYDataGrid()
        {
            ABBYYDataModel model = new ABBYYDataModel()
            {
                ControlNumber = "1191167"
            };
            model.FillABBYYDataGrid();
            var row = model.ABBYYData.Rows;
            var col = model.ABBYYData.Columns;
            Assert.AreEqual("Farm to Market Road 482", row[0][col[4]]);
        }
    }
}
