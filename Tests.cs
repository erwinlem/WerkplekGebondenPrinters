using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace WerkplekGebondenPrinter {
    [TestClass]
    public class Tests {
        [TestMethod]
        public void TestAD() {
            Config c = new Config();
            Assert.AreNotEqual(0, c.GetPrintersAD().Count);
        }
    }
}
