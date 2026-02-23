using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;

namespace WerkplekGebondenPrinter {

/*
wel even de volgende tabellen aanmaken, getest met sqlserver :

create table WerkplekPrinters (
	id INT IDENTITY(1, 1),
	werkplek nvarchar(max),
	printer nvarchar(max)
)

*/

    public class ConfigLoaderSQL : IConfigLoader {
        string werkplek;
        public string Werkplek { get => werkplek; set => werkplek = value; }
        List<string> printers = new List<string>();
        public List<string> Printers { get => printers; set => printers = value; }

        public void LoadPrinters() {
            DataSet dataset = new DataSet();
            using (SqlConnection connection = new SqlConnection(Config.Settings["ConfigLoaderSQL.connectionString"])) {
                SqlDataAdapter adapter = new SqlDataAdapter();
#warning sql inject
                adapter.SelectCommand = new SqlCommand("select printer from WerkplekPrinters where werkplek = '"+System.Environment.MachineName+"'", connection);
                adapter.Fill(dataset);
                foreach (DataRow row in dataset.Tables[0].Rows) {
                    printers.Add((string)row[0]);
                }
            }
        }

        public void SavePrinters() {
            // todo: upsert? anders gaan die id-tjes ook zo hard.
            using (SqlConnection connection = new SqlConnection(Config.Settings["ConfigLoaderSQL.connectionString"])) {

                connection.Open();
                SqlTransaction transaction = connection.BeginTransaction();
                try {
                    // haal oude weg
                    using (SqlCommand cmd = new SqlCommand(
                        "delete from WerkplekPrinters where werkplek = @werkplek", connection, transaction)) {
                        cmd.Parameters.AddWithValue("@werkplek", System.Environment.MachineName);
                        cmd.ExecuteNonQuery();
                    }

                    // stop nieuwe erin
                    foreach (var p in printers) {
                        using (SqlCommand cmd = new SqlCommand(
                            "insert into WerkplekPrinters(werkplek, printer) values (@werkplek, @printer) ", connection, transaction)) {
                            cmd.Parameters.AddWithValue("@werkplek", System.Environment.MachineName);
                            cmd.Parameters.AddWithValue("@printer", p);
                            cmd.ExecuteNonQuery();
                        }
                    }

                    // Commit if all succeed
                    transaction.Commit();
                    Trace.TraceInformation("Transaction committed successfully.");
                } catch (Exception ex) {
                    // Rollback if any fail
                    transaction.Rollback();
                    Trace.TraceError("Transaction rolled back due to error: " + ex.Message);
                }
            }
        }
    }
}
