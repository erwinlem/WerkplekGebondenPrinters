using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace WerkplekGebondenPrinter {
    [TestClass]
    public class Tests {
        private TestContext testContextInstance;

        [TestMethod]
        public void Test_ConfigLoaderBestand_load() {
            IConfigLoader cl = new ConfigLoaderBestand();
            cl.LoadPrinters();
            foreach (string s in cl.Printers) { 
                testContextInstance.WriteLine("printer:"+s);
            }
        }

        [TestMethod]
        public void Test_AD() {
            Config c = new Config();
            Assert.AreNotEqual(0, c.GetPrintersAD().Count);
        }
    }
}
