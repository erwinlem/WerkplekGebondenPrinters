using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace WerkplekGebondenPrinter {
    internal class PrinterLoaderAD : IPrinterLoader {
        List<PrinterInfo> IPrinterLoader.LoadPrinters() {
            var printers = new List<PrinterInfo>();
            using (var searcher = new DirectorySearcher("(objectCategory=printQueue)")) {
                searcher.PropertiesToLoad.Add("printername");
                searcher.PropertiesToLoad.Add("description");
                searcher.PropertiesToLoad.Add("location");
                searcher.PropertiesToLoad.Add("UncName");


                foreach (SearchResult result in searcher.FindAll()) {
                    var desc = "";
                    if (result.Properties.Contains("description")) {
                        desc = result.Properties["description"]?[0]?.ToString();
                    }
                    // TODO: Config direct aanspreken is niet zo netjes
                    if (Regex.IsMatch(desc, Config.FilterRegex, RegexOptions.IgnoreCase)) {
                        var Location = "";
                        if (result.Properties.Contains("location")) {
                            Location = result.Properties["location"]?[0]?.ToString();
                        }

                        printers.Add(new PrinterInfo {
                            PrinterName = result.Properties["printername"]?[0]?.ToString(),
                            Description = desc,
                            Location = Location,
                            UncName = result.Properties["UncName"]?[0]?.ToString()
                        });
                    }
                }
            }
            return printers;
    }
}
}
