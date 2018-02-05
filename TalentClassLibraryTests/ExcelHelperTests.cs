using Microsoft.VisualStudio.TestTools.UnitTesting;
using TalentClassLibrary;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TalentClassLibrary.Tests
{
    [TestClass()]
    public class ExcelHelperTests
    {
        [TestMethod()]
        public void ExportMultipleContactSituationTest()
        {
            var input = TestData.GetInstance().TestExcelContactSituationData();
            var expected = "匯出成功";
            var actual = ExcelHelper.GetInstance().ExportMultipleContactSituation(input, @"..\..\..\Template");
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void ExportInterviewDataTest()
        {
            var input = TestData.GetInstance().TestExcelInterviewInfoData();
            var input1 = TestData.GetInstance().TestExcelProjectExperienceData();
            var input2 = TestData.GetInstance().TestExcelInterviewResultData();
            var expected = "匯出成功";
            //var actual = ExcelHelper.GetInstance().ExportInterviewData(input, input1, input2);
            // Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void ExportAllDataTest()
        {
            var input = TestData.GetInstance().TestExcelContactSituationData();
            var input1 = TestData.GetInstance().TestExcelInterviewInfoData();
            var input2 = TestData.GetInstance().TestExcelProjectExperienceData();
            var expected = "匯出成功";
            //  var actual = ExcelHelper.GetInstance().ExportAllData(input, input1,input2);
            // Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void ImportNewTalentTest()
        {
            var input = @"..\..\..\Template\匯入資料範本(google duc).xlsx";
            ExcelHelper.GetInstance().ImportNewTalent(input);
        }

        [TestMethod()]
        public void ImportOldTalentTest()
        {
            var input = @"..\..\..\Template\匯入資料範本(人資系統).xlsx";
            ExcelHelper.GetInstance().ImportOldTalent(input);
        }

        [TestMethod()]
        public void ImportInterviewDataTest()
        {
            var input = @"..\..\..\Template\匯入面談資料範本.xlsx";          
            var actual = ExcelHelper.GetInstance().ImportInterviewData(input);          
        }
    }
}