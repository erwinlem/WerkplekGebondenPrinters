using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

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
        string connectionString = "Server=SQLSERVER;Database=DATABASE;Trusted_Connection=True;";
        string werkplek;
        public string Werkplek { get => werkplek; set => werkplek = value; }
        List<string> printers = new List<string>();
        public List<string> Printers { get => printers; set => printers = value; }

        public void LoadPrinters() {
            DataSet dataset = new DataSet();
            using (SqlConnection connection = new SqlConnection(connectionString)) {
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
            throw new NotImplementedException();
        }
    }
}
