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
            ABBYYDataModel testModel = new ABBYYDataModel();
            testModel.ControlNumber = null;
            testModel.FillABBYYDataGrid();
        }

        [TestMethod]
        public void TestFillABBYYDataGrid_ReturnIfEmptyString()
        {
            ABBYYDataModel testModel = new ABBYYDataModel();
            testModel.ControlNumber = "";
            testModel.FillABBYYDataGrid();
        }
    }
}
