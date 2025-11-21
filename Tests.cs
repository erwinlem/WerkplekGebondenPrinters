using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Diagnostics;

namespace WerkplekGebondenPrinter {
    [TestClass]
    public class Tests {
        [TestMethod]
        public void Test_ConfigLoaderBestand_load() {
            IConfigLoader cl = new ConfigLoaderBestand();
            cl.LoadPrinters();
            foreach (string s in cl.Printers) {
                Trace.WriteLine("printer:"+s);
            }
        }

        [TestMethod]
        public void Test_ConfigLoaderSQL_load() {
            IConfigLoader cl = new ConfigLoaderSQL();
            cl.LoadPrinters();
            foreach (string s in cl.Printers) {
                Trace.WriteLine("printer:" + s);
            }
        }

        [TestMethod]
        public void Test_AD() {
            Config c = new Config();
            Assert.AreNotEqual(0, c.GetPrintersAD().Count);
        }
    }
}
