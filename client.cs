using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.DirectoryServices;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Management;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using System.Windows.Media.Imaging;
using System.Xml;
using Application = System.Windows.Application;
using Button = System.Windows.Controls.Button;
using DataGrid = System.Windows.Controls.DataGrid;
using Image = System.Windows.Controls.Image;
using TextBox = System.Windows.Controls.TextBox;

namespace WerkplekGebondenPrinter {
    public struct PrinterInfo {
        public string PrinterName;
        public string Description;
        public string Location;
        public string UncName;
    }

    public class Config {
        private const string FilterRegex = "^(ETK-Medicatie|ETK-LAB|ETK-RODEBALK|ETK-STAM|A4|ETK-Patient|Polsband|PolsbandB|PolsbandK)$";

        public List<PrinterInfo> GetPrintersAD() {
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
                    if (Regex.IsMatch(desc, FilterRegex)) {
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

        public class WindowData {
            public string TB_Werkplek { get; set; } = "Werkplek - " + Environment.GetEnvironmentVariable("CLIENTNAME");
            private string _TB_Filter;
            public string TB_Filter {
                get => _TB_Filter;
                set {
                    _TB_Filter = value;
                }
            }

            public void HandleButtonClick(string name, MainWindow window) {
                switch (name) {
                    case "butt_reset": window.LoadPrinters(); break;
                    case "butt_cancel": window.Close(); break;
                    case "butt_set":
                        var s = ((DataRowView)((DataGrid)window.FindName("printers")).SelectedItem).Row[0];
                        window.setPrinter(s.ToString());
                        break;
                    case "butt_unset":
                        var type = ((DataRowView)((DataGrid)window.FindName("printerType")).SelectedItem).Row[0];
                        window.clearPrinter(type.ToString());
                        break;
                    case "butt_ok":
                        window.ApplyChanges();
                        break;
                }
            }
        }

        public class MainWindow : System.Windows.Window {
            private WindowData viewModel = new WindowData();
            private List<PrinterInfo> printers;
            private DataTable printerTable = new DataTable();
            private DataTable printerTypeTable = new DataTable();
            private string selectedType = "";
            Config config = new Config();

            public MainWindow() {
                DataContext = viewModel;

                printerTable.Columns.Add("PrinterName");
                printerTable.Columns.Add("Description");
                printerTable.Columns.Add("Location");

                printerTypeTable.Columns.Add("Type");
                printerTypeTable.Columns.Add("Printer");
                printerTypeTable.Columns.Add("Location");

                LoadPrinters();
                SetupUI();
                this.Title = "Werkplekgebondenprinter 2.0";
            }

            public void clearPrinter(string type) {
                DataRow[] pt = printerTypeTable.Select("Type = '" + type + "'");
                if (pt.Length == 0) {
                    return; // printer type niet gevonden
                }
                pt[0][1] = "";
                pt[0][2] = "";

            }

            public void setPrinter(string p) {
                Trace.TraceInformation("Installed Printer: " + p);
                DataRow[] p2 = printerTable.Select("PrinterName = '" + p + "'");
                if (p2.Length == 0) {
                    return; // printer niet gevonden in ad
                }
                        ;
                Trace.TraceInformation("type: " + p2[0][1]);
                DataRow[] pt = printerTypeTable.Select("Type = '" + p2[0][1] + "'");
                if (pt.Length == 0) {
                    return; // printer type niet gevonden
                }
                pt[0][1] = p;
                pt[0][2] = p2[0][2];
            }

            public void LoadPrinters() {

                // active directory
                printers = config.GetPrintersAD();

                printerTable.Clear();

                foreach (var p in printers)
                    printerTable.Rows.Add(p.PrinterName, p.Description, p.Location);

                // fill printertypes
                var tmp = printers.GroupBy(x => x.Description)
                          .Select(y => new {
                              Type = y.Key,
                              Printer = "",
                              Location = "",
                          });

                printerTypeTable.Clear();

                foreach (var p in tmp) {
                    printerTypeTable.Rows.Add(p.Type, p.Printer, p.Location);
                    Trace.WriteLine(p);
                }

                // computer printers
                try {
                    foreach (string printer in PrinterSettings.InstalledPrinters) {
                        setPrinter(printer.Split('\\').Last());
                    }
                } catch (Win32Exception e) {
                    // printspooler waarschijk disabled, TODO: vullen met dummy

                }
            }

        #region add/remove printer
        public void AddPrinter(string printerPath) {
                try {
                    Trace.TraceInformation($"Adding Printer {printerPath}");
                    var managementClass = new ManagementClass("Win32_Printer");
                    var inputParams = managementClass.GetMethodParameters("AddPrinterConnection");
                    inputParams["Name"] = printerPath;

                    managementClass.InvokeMethod("AddPrinterConnection", inputParams, null);
                    
                } catch (Exception ex) {
                    Trace.TraceError($"Failed to add printer {printerPath}: {ex.Message}");
                }
            }

            public void RemovePrinter(string printerName) {
                try {
                    string query = $"SELECT * FROM Win32_Printer WHERE ShareName = '{printerName.Replace("\\", "\\\\")}'";
                    using (var searcher = new ManagementObjectSearcher(query)) {
                        foreach (ManagementObject printer in searcher.Get()) {
                            printer.Delete();
                            Trace.TraceInformation($"Printer removed: {printerName}");
                        }
                    }
                } catch (Exception ex) {
                    Trace.TraceInformation($"Failed to remove printer {printerName}: {ex.Message}");
                }
            }
        #endregion

        string xaml = @"
    <Grid xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
      xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml' Background='{DynamicResource Theme_BackgroundBrush}'>
        <TextBlock Name='TB_Werkplek' Text='{Binding TB_Werkplek}' FontSize='20' Margin='8,8,0,8' HorizontalAlignment='Left' VerticalAlignment='Top' Height='32' Width='624' />
        <TextBlock Text='Filter:' FontSize='14' Margin='648,16,0,0' HorizontalAlignment='Left' VerticalAlignment='Top' Height='24' />
        <TextBox Name='TB_Filter' Text='{Binding TB_Filter, UpdateSourceTrigger=PropertyChanged}' Margin='720,8,4,8' FontSize='24' VerticalAlignment='Top' Height='32' IsEnabled='True' />
        <Button Name='butt_set' Content='← GEBRUIK' Width='70' Height='40' Margin='640,108,4,4' VerticalAlignment='Top' HorizontalAlignment='Left' />
        <Button Name='butt_reset' Content='↺ Herstel' Width='70' Height='40' Margin='640,156,4,4' VerticalAlignment='Top' HorizontalAlignment='Left' />
        <Button Name='butt_unset' Content='❌ Verwijder' Width='70' Height='40' Margin='640,206,4,4' VerticalAlignment='Top' HorizontalAlignment='Left' />
        <DataGrid Name='printerType' Margin='8,48,0,328' HorizontalAlignment='Left' Width='624' IsReadOnly='True' SelectionMode='Single' CanUserSortColumns='True' />
        <DataGrid Name='printers' Margin='720,48,8,48' IsReadOnly='True' SelectionMode='Single' />
        <StackPanel Orientation='Horizontal' HorizontalAlignment='Right' VerticalAlignment='Bottom' Margin='4,4,4,4'>
            <Button Name='butt_cancel' Content='Annuleren' Width='100' Margin='4,4,4,4' />
            <Button Name='butt_ok' Content='Toepassen' Margin='4,4,4,4' Width='100' />
        </StackPanel>
        <Image Name='img_voorbeeld' Margin='8,0,0,8' HorizontalAlignment='Left' VerticalAlignment='Bottom' Width='320' Height='320' />
    </Grid>
";

            private void SetupUI() {
                var reader = XmlReader.Create(new StringReader(xaml));
                var content = (Grid)XamlReader.Load(reader);
                this.Content = content;
                NameScope.SetNameScope(this, NameScope.GetNameScope(content));
                this.DataContext = viewModel;

                // write up the buttons
                foreach (var name in new[] { "set", "unset", "ok", "reset", "cancel" }) {
                    var btn = (Button)(this.FindName("butt_" + name));
                    btn.Click += (s, e) => viewModel.HandleButtonClick("butt_" + name, this);
                }

                var img = (Image)this.FindName("img_voorbeeld");
                img.Source = LoadImage("voorbeeld/kat.png");

                ((TextBox)this.FindName("TB_Filter")).TextChanged += (s, e) => Refresh();

                var grid = ((DataGrid)this.FindName("printerType"));
                grid.SelectionChanged += (s, e) => {
                    if (grid.SelectedItem is DataRowView row) {
                        selectedType = row[0].ToString(); // FIXME: "Description"
                        string path = "voorbeeld/" + selectedType + ".png";
                        if (!File.Exists(path)) path = "voorbeeld/kat.png";
                        img.Source = LoadImage(path);
                        Refresh();
                    }
                };

                ((DataGrid)this.FindName("printers")).ItemsSource = printerTable.DefaultView;
                ((DataGrid)this.FindName("printerType")).ItemsSource = printerTypeTable.DefaultView;
            }

            public void ApplyChanges() {
                this.Cursor = System.Windows.Input.Cursors.Wait;

                // Get new printers from printerType DataGrid
                var printerTypeGrid = (DataGrid)this.FindName("printerType");
                var printersNew = printerTypeGrid.Items
                    .Cast<dynamic>()
                    .Where(item => !string.IsNullOrEmpty((string)item.Row[1]))
                    .Select(item => (string)item.Row[1])
                    .Select(item => printers.Where(x => (x.PrinterName == item)).ToArray()[0].UncName)
                    .ToList();

                // Get old printers from printersCurrent
                List<string> printersOld;

                try {
                    printersOld = PrinterSettings.InstalledPrinters
                        .Cast<string>()
                        .Where(item => printers.Where(x => (x.UncName == item)).ToList().Count != 0) // filter niet gepubliceerde printers
                        .ToList();
                } catch (Win32Exception e) {
                    // printspooler waarschijk disabled, TODO: vullen met dummy
                    return;

                }

                Trace.TraceInformation("nieuwe printers : " + string.Join(",", printersNew));
                Trace.TraceInformation("oude printers : " + string.Join(",", printersOld));

                var compare = printersNew
                    .Union(printersOld)
                    .Distinct()
                    .Select(p => new {
                        Printer = p,
                        Side = printersNew.Contains(p) && printersOld.Contains(p) ? "==" :
                               printersNew.Contains(p) ? "<=" :
                               printersOld.Contains(p) ? "=>" : "?"
                    });

                foreach (var entry in compare) {
                    var printer = entry.Printer;

                    Trace.TraceInformation("printer " + entry.Printer);

                    switch (entry.Side) {
                        case "==":
                            Trace.TraceInformation("keeping " + printer);
                            break;
                        case "<=":
                            Trace.TraceInformation("printer " + printer + " toevoegen");
                            AddPrinter(printer);
                            Trace.TraceInformation("printer " + printer + " toegevoegd");
                            break;
                        case "=>":
                            Trace.TraceInformation("weg printer " + printer);
                            RemovePrinter(printer);
                            Trace.TraceInformation("weggehaald printer " + printer);
                            break;
                        default:
                            Trace.TraceWarning("Geen idee wat dit is " + printer);
                            break;
                    }
                }

                LoadPrinters();
                this.Cursor = System.Windows.Input.Cursors.Arrow;

                // Save printer list to file
                var outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "computers", Environment.GetEnvironmentVariable("CLIENTNAME") + ".txt");
                var finalList = printersNew
                    .Select(p => printers.FirstOrDefault(x => x.PrinterName == p).Location?.Replace(".intranet.local", ""))
                    .Where(x => !string.IsNullOrEmpty(x))
                    .OrderBy(x => x);

                File.WriteAllLines(outputPath, finalList);

                Trace.TraceInformation("KLAAR!");
            }

            public void Refresh() {
                string filter = "(PrinterName LIKE '%" + viewModel.TB_Filter + "%' OR Location LIKE '%" + viewModel.TB_Filter + "%')";
                if (!string.IsNullOrEmpty(selectedType))
                    filter += " AND Description LIKE '" + selectedType + "'";
                printerTable.DefaultView.RowFilter = filter;
            }

            public static BitmapImage LoadImage(string path) {
                using (var bmp = new Bitmap(path))
                using (var ms = new MemoryStream()) {
                    bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    ms.Position = 0;
                    var img = new BitmapImage();
                    img.BeginInit();
                    img.StreamSource = ms;
                    img.CacheOption = BitmapCacheOption.OnLoad;
                    img.EndInit();
                    img.Freeze();
                    return img;
                }
            }
        }

        public class App : Application {

        [STAThread]
        public static int Main(string[] args) {
            Trace.AutoFlush = true;
            
            Trace.TraceInformation("Werkplekgebonenprinter wordt opgestart");

            for (int i = 0; i < args.Length; i++) {
                switch (args[i]) {
                    case "-d":
                    case "--debug":
                        Trace.TraceInformation("Debug logging is AAN");
                        break;
                    case "-l":
                        // TODO: voorkomen dat het op 2 regels komt
                        var li = new TextWriterTraceListener(args[++i]);
                        li.TraceOutputOptions = TraceOptions.DateTime;
                        Trace.Listeners.Add(li);
                        break;
                    case "-v":
                    default:
                        Trace.TraceError($"Unknown argument: {args[i]}");
                        return 1;
                }
            }

            var app = new App();
            var window = new MainWindow();
            app.Run(window);
            Trace.WriteLine("Werkplekgebonenprinter wordt afgesloten");
            return 0;
        }
    }
}
