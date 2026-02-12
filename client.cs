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
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
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

    // interface voor het laden van de printer/werkplek mapping & overige opties
    // interface is nodig vanwege de verschillende backends (files/db/etc)
    public interface IConfigLoader {
        string Werkplek { get; set; }
        List<string> Printers { get; set; }
        void LoadPrinters();
        void SavePrinters();
    }

    // haal een lijst op van alle mogelijke printers
    public interface IPrinterLoader {
        List<PrinterInfo> LoadPrinters();
    }

    public class Config {
        public static Config config = new Config();
        public static string CommentFilter = "^(ETK-Medicatie|ETK-LAB|ETK-RODEBALK|ETK-STAM|A4|ETK-Patient|Polsband|PolsbandB|PolsbandK|ETK-PAPO)$";
        public IConfigLoader werkplekPrinter = new ConfigLoaderBestand(); // voor nu alleen nog even bestanden
        public IPrinterLoader printerLoader = new PrinterLoaderDummy();

        private List<PrinterInfo> printers;
        public List<PrinterInfo> Printers {
            get {
                if (printers == null) {
                    printers = printerLoader.LoadPrinters();
                }
                return printers;
            }
        }

        public List<string> GetInstalledWPGPrinters() {
            try {
                Trace.TraceInformation(String.Join("\n",
                        PrinterSettings.InstalledPrinters.Cast<string>()
                    .Select(item => Regex.Replace(item, @"^\\\\([^\.]*)\\", @"\\$1." + System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName + "\\"))
                    ));
                return PrinterSettings.InstalledPrinters
                    .Cast<string>()
                    // Zorg ervoor dat er altijd een FQDN wordt gebruikt bij het vergelijken, zo komen ze ook uit het ad rollen.
                    // \\printserver1\print1 wordt dan \\printserver1.domein.local\print1
                    .Where(item => Printers.Where(x => (x.UncName.Equals(Regex.Replace(item, @"\\\\([^\.]*)\\", @"\\$1." + System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName + "\\"),StringComparison.OrdinalIgnoreCase))).ToList().Count != 0) // filter niet gepubliceerde printers
                    .ToList();
            } catch (Win32Exception e) {
                Trace.TraceError("fout printers ophalen, spooler disabled?");
                return new List<string> { "error" } ;

            }
        }

        public void SaveConfig(List<string> printersNew) {
            werkplekPrinter.Printers = printersNew;
            werkplekPrinter.SavePrinters();
        }

        public void ApplyConfig() {
            werkplekPrinter.LoadPrinters();
            ApplyPrintlist(werkplekPrinter.Printers, GetInstalledWPGPrinters());
        }

        public void ApplyPrintlist(List<string> printersNew, List<string> printersOld) {
            Trace.TraceInformation("nieuwe printers : " + string.Join(",", printersNew));
            Trace.TraceInformation("oude printers : " + string.Join(",", printersOld));

            var compare = printersNew
                .Union(printersOld, StringComparer.OrdinalIgnoreCase)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Select(p => new {
                    Printer = p,
                    Side = printersNew.Contains(p, StringComparer.OrdinalIgnoreCase) && printersOld.Contains(p, StringComparer.OrdinalIgnoreCase) ? "==" :
                           printersNew.Contains(p, StringComparer.OrdinalIgnoreCase) ? "<=" :
                           printersOld.Contains(p, StringComparer.OrdinalIgnoreCase) ? "=>" : "?"
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
                        Config.AddPrinter(printer);
                        Trace.TraceInformation("printer " + printer + " toegevoegd");
                        break;
                    case "=>":
                        Trace.TraceInformation("weg printer " + printer);
                        Config.RemovePrinter(printer);
                        Trace.TraceInformation("weggehaald printer " + printer);
                        break;
                    default:
                        Trace.TraceWarning("Geen idee wat dit is " + printer);
                        break;
                }
            }

            // notificatie popup
            string message = string.Join("\n",
                    compare.Where(item => item.Side != "==")
                        .Select(p => (p.Side == "=>" ? "❌" : "✅") + (p.Printer.Split('\\').Last()))
                    );
            if (message == "") {
                // geen aanpassing, dus geen popup nodig
                return;
            }
            NotifyIcon notifyIcon1 = new NotifyIcon();
            notifyIcon1.Visible = true;
            notifyIcon1.Icon = SystemIcons.Information;
            notifyIcon1.BalloonTipTitle = "Printers zijn aangepast";
            notifyIcon1.BalloonTipText = message;
            notifyIcon1.BalloonTipIcon = ToolTipIcon.None;
            notifyIcon1.ShowBalloonTip(5000);
        }

        #region add/remove printer
        public static void AddPrinter(string printerPath) {
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

        public static void RemovePrinter(string printerName) {
            try {
                string query = $"SELECT * FROM Win32_Printer WHERE Name = '{printerName.Replace("\\", "\\\\")}'";
                using (var searcher = new ManagementObjectSearcher(query)) {
                    foreach (ManagementObject printer in searcher.Get()) {
                        Trace.TraceInformation($"removing: {printerName}");
                        printer.Delete();
                        Trace.TraceInformation($"Printer removed: {printerName}");
                    }
                }
            } catch (Exception ex) {
                Trace.TraceInformation($"Failed to remove printer {printerName}: {ex.Message}");
            }
        }
        #endregion

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
                case "butt_reset":
                    window.LoadPrinters();
                    break;
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
        private DataTable printerTable = new DataTable();
        private DataTable printerTypeTable = new DataTable();
        private string selectedType = "";

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
            printerTable.Clear();

            foreach (var p in Config.config.Printers)
                printerTable.Rows.Add(p.PrinterName, p.Description, p.Location);

            // fill printertypes
            var tmp = Config.config.Printers.GroupBy(x => x.Description)
                      .Select(y => new {
                          Type = y.Key,
                          Printer = "",
                          Location = "",
                      })
                      .OrderBy(x => x.Type);

            printerTypeTable.Clear();

            foreach (var p in tmp) {
                printerTypeTable.Rows.Add(p.Type, p.Printer, p.Location);
                Trace.WriteLine(p);
            }

            // computer printers
            try {
                foreach (string printer in Config.config.GetInstalledWPGPrinters()) {
                    setPrinter(printer.Split('\\').Last());
                }
            } catch (Win32Exception e) {
                // printspooler waarschijk disabled, TODO: vullen met dummy

            }
        }

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
            img.Source = LoadImage("voorbeeld/onbekend.png");

            ((TextBox)this.FindName("TB_Filter")).TextChanged += (s, e) => Refresh();

            var grid = ((DataGrid)this.FindName("printerType"));
            grid.SelectionChanged += (s, e) => {
                if (grid.SelectedItem is DataRowView row) {
                    selectedType = row[0].ToString(); // FIXME: "Description"
                    string path = "voorbeeld/" + selectedType + ".png";
                    if (!File.Exists(path)) path = "voorbeeld/onbekend.png";
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
                .Select(item => Config.config.Printers.Where(x => (x.PrinterName == item)).ToArray()[0].UncName)
                .ToList();

            // eerst config toepassen, dan pas opslaan
            // dit voorkomt dat als er een foute printerdriver/exception is de oude situatie blijft
            Config.config.ApplyPrintlist(printersNew, Config.config.GetInstalledWPGPrinters());
            LoadPrinters();
            Config.config.SaveConfig(printersNew);
            this.Cursor = System.Windows.Input.Cursors.Arrow;

            Trace.TraceInformation("KLAAR!");
        }

        public void Refresh() {
            string filter = "(PrinterName LIKE '%" + viewModel.TB_Filter + "%' OR Location LIKE '%" + viewModel.TB_Filter + "%')";
            if (!string.IsNullOrEmpty(selectedType))
                filter += " AND Description LIKE '" + selectedType + "'";
            printerTable.DefaultView.RowFilter = filter;
        }

        public static BitmapImage LoadImage(string path) {
            // het moet via een memorystream omdat het bestand anders gelocked wordt door xaml
            // gewoon direct toekennen zou een stuk makkelijker zijn
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
            bool sync = false;

            Trace.TraceInformation("Werkplekgebonenprinter wordt opgestart");

            for (int i = 0; i < args.Length; i++) {
                switch (args[i]) {
                    case "--PrinterLoader":
                        switch (args[++i]) {
                            case "PrinterLoaderAD":
                                Config.config.printerLoader = new PrinterLoaderAD();
                                break;
                            case "PrinterLoaderDummy":
                                Config.config.printerLoader = new PrinterLoaderDummy();
                                break;
                            default:
                                Trace.TraceError($"onbekende loader: {args[i]}");
                                return 1;
                        }
                        break;
                    case "--CommentFilter":
                        Config.CommentFilter = args[++i];
                        break;
   
                    case "-d":
                    case "--debug":
                        Trace.TraceInformation("Debug logging is AAN");
                        break;
                    case "-s":
                    case "--sync":
                        Trace.TraceInformation("printers worden gesynced");
                        sync = true;
                        break;
                    case "--cwd":
                        // in ivanti taak kan je niet makkelijk de current working directory instellen
                        Directory.SetCurrentDirectory(args[++i]);
                        break;
                    case "-l":
                        // TODO: voorkomen dat het op 2 regels komt
                        var li = new TextWriterTraceListener(args[++i]);
                        li.TraceOutputOptions = TraceOptions.DateTime;
                        Trace.Listeners.Add(li);
                        break;
                    default:
                        Trace.TraceError($"Unknown argument: {args[i]}");
                        return 1;
                }
            }

            try {
                if (sync) {
                    Config.config.ApplyConfig();
                } else { 
                    var app = new App();
                    var window = new MainWindow();
                    app.Run(window);
                }
            } catch ( Exception ex ) {
                Trace.TraceError($"Algemene fout: {ex.Message}");
                Trace.TraceError($"stacktrace: {ex.StackTrace}");
            }
            Trace.WriteLine("Werkplekgebonenprinter wordt afgesloten");
            return 0;
        }
    }
}
