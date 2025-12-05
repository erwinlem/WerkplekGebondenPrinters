using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WerkplekGebondenPrinter {
    internal class PrinterLoaderDummy : IPrinterLoader {
        List<PrinterInfo> IPrinterLoader.LoadPrinters() {
            var l = new List<PrinterInfo>();

            string[] printers = {
                    "prn001|A4|cardio ❤️ 1e|\\\\printer01\\prn001",
                    "prn002|A4|cardio 2e|\\\\printer01\\prn002",
                    "prn003|A4|endo 1e|\\\\printer01\\prn003",
                    "prn004|A4|endo balie|\\\\printer01\\prn004",
                    "prn005|A4|cardio 1e verdieping|\\\\printer01\\prn005",
                    "prn006|label|endo 1e|\\\\printer01\\prn006",
                    "prn007|label|endo balie|\\\\printer01\\prn007",
                    "prn008|label|cardio 1e verdieping|\\\\printer01\\prn008",
                    "prn009|recept|Apotheek 1e verdieping|\\\\printer01\\prn009",
                    "prn010|recept|Apotheek 2e verdieping|\\\\printer01\\prn010",
                    "prn011|rode balk|centrale apotheek 💊|\\\\printer01\\prn011",
                    "prn012|rode balk|Apotheek 2e verdieping|\\\\printer01\\prn012",
                    "prn013|rode balk|centrale apotheek 💊|\\\\printer01\\prn013"
            };

            foreach (var printer in printers) {
                var p2 = printer.Split('|');
                l.Add(new PrinterInfo { PrinterName = p2[0], Description = p2[1], Location = p2[2], UncName = p2[3] });
            }

            return l;
        }
    }
}
