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
    public class TalentSearchTests
    {
        [TestMethod()]
        public void SelectIdByFilterTest()
        {
            TalentSearch.GetInstance().SelectIdByFilter(string.Empty, string.Empty, "C#,JAVA", string.Empty, string.Empty, "2017-12-15", "2018-01-05", "不限", "不限", string.Empty, string.Empty);
        }
    }
}