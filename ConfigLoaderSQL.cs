using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Reflection;

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
            printers.Clear();
            DataSet dataset = new DataSet();
            using (SqlConnection connection = new SqlConnection(Config.Settings["ConfigLoaderSQL.connectionString"])) {
                SqlDataAdapter adapter = new SqlDataAdapter();
                adapter.SelectCommand = new SqlCommand("select printer from WerkplekPrinters where werkplek = @werkplek", connection);
                adapter.SelectCommand.Parameters.AddWithValue("@werkplek", Config.Hostname);
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
                        cmd.Parameters.AddWithValue("@werkplek", Config.Hostname);
                        cmd.ExecuteNonQuery();
                    }

                    // stop nieuwe erin
                    foreach (var p in printers) {
                        using (SqlCommand cmd = new SqlCommand(
                            "insert into WerkplekPrinters(werkplek, printer) values (@werkplek, @printer) ", connection, transaction)) {
                            cmd.Parameters.AddWithValue("@werkplek", Config.Hostname);
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

    /*
    -- ja, ik weet het, max is slecht
CREATE TABLE[Logs] (
        [id] INT IDENTITY(1,1) PRIMARY KEY,
    [timestamp] DATETIME DEFAULT GETDATE(),
    [message] NVARCHAR(max),
    [category] NVARCHAR(max) NULL,
	[machineName] NVARCHAR(max) NULL,
    [userName] NVARCHAR(max) NULL
   );
    */
    public class SqlTraceListener : TraceListener {
        public override void Write(string message) => LogToDatabase(message, null);
        public override void WriteLine(string message) => LogToDatabase(message, null);

        public override void TraceEvent(TraceEventCache eventCache, string source, TraceEventType eventType, int id, string message) {
            LogToDatabase(message, eventType.ToString());
        }

        private void LogToDatabase(string message, string category) {
            if (string.IsNullOrWhiteSpace(message)) return;

            const string query = @"
            INSERT INTO Logs (Message, Category, MachineName, UserName) 
            VALUES (@Message, @Category, @MachineName, @UserName)";

            try {
                using (var connection = new SqlConnection(Config.Settings["ConfigLoaderSQL.connectionString"])) {
                    connection.Open();
                    using (var command = new SqlCommand(query, connection)) {
                        command.Parameters.AddWithValue("@Message", message);
                        command.Parameters.AddWithValue("@Category", (object)category ?? DBNull.Value);
                        command.Parameters.AddWithValue("@MachineName", Config.Hostname);
                        command.Parameters.AddWithValue("@UserName", Environment.UserName);

                        command.ExecuteNonQuery();
                    }
                }
            } catch {
                // Fail-safe
            }
        }
    }
}
