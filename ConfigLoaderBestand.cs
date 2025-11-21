using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WerkplekGebondenPrinter {
    // simpele dump per machine in een bestand
    public class ConfigLoaderBestand : IConfigLoader {
        string werkplek;
        string IConfigLoader.Werkplek { get => werkplek; set => werkplek = value; }
        List<string> printers = new List<string>();
        List<string> IConfigLoader.Printers { get => printers; set => printers = value; }

        public void LoadPrinters() {
            var filename = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "computers", Environment.GetEnvironmentVariable("CLIENTNAME") + ".txt");
            try {
                printers = new List<string>(File.ReadAllLines(filename));
            } catch (Exception e) {
                Trace.TraceError($"error reading file {filename}");
                Trace.TraceError($"exception {e}");
            }
        }

        public void SavePrinters() {
            // Save printer list to file
            var outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "computers", Environment.GetEnvironmentVariable("CLIENTNAME") + ".txt");
            File.WriteAllLines(outputPath, printers);
        }
    }
}
