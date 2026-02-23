using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Configuration;
using System.Diagnostics;
using System.Linq;

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
            App.ParseArguments(ConfigurationManager.AppSettings);
            IConfigLoader cl = new ConfigLoaderSQL();
            cl.LoadPrinters();
            foreach (string s in cl.Printers) {
                Trace.WriteLine("printer:" + s);
            }
        }

        [TestMethod]
        public void Test_ConfigLoaderSQL_Save() {
            App.ParseArguments(ConfigurationManager.AppSettings);
            IConfigLoader cl = new ConfigLoaderSQL();
            cl.LoadPrinters();
            cl.Printers.Clear();
            cl.Printers.Add("\\\\printer\\testprinter-static");
            cl.Printers.Add("\\\\printer\\testprinter-"+(new DateTime()).ToString("yyyy-MM-dd-h-mm-tt"));
            cl.SavePrinters();
        }

        [TestMethod]
        public void Test_AD() {
            Config c = new Config();
            Assert.AreNotEqual(0, c.Printers.Count);
        }

        [TestMethod]
        public void Test_Remove() {
            Config c = new Config();
            var pd = c.DiffPrinters(
                    new System.Collections.Generic.List<string>() { @"\\print01\a", @"\\print01\b" },
                    new System.Collections.Generic.List<string>() { @"\\print01\b", @"\\print01\c" }
                );
            foreach (var p in pd) {
                Trace.WriteLine("printer:" + p);
            }
        }
        [TestMethod]
        public void Test_RemoveDomein() {
            Config c = new Config();
            var pd = c.DiffPrinters(
                    new System.Collections.Generic.List<string>() { @"\\print01.domein\a", @"\\print01.domein\b" },
                    new System.Collections.Generic.List<string>() { @"\\print01\b", @"\\print01\c" }
                );
            foreach (var p in pd) {
                Trace.WriteLine("printer:" + p);
            }
        }
    }
}
