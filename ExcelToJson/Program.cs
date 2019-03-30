using System.Collections.Generic;
using System.Data.Common;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using ExcelDataReader;
using Newtonsoft.Json;

namespace ExcelToJson
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var inFilePath = args[0];
            var outFilePath = args[1];
            var sheetName = args[2];
            UsingExcelDataReader(inFilePath, outFilePath);
        }

        private static void UsingOleDb(string inFilePath, string outFilePath, string sheetName)
        {
            //"HDR=Yes;" indicates that the first row contains column names, not data.
            var connectionString = $@"
                Provider=Microsoft.ACE.OLEDB.12.0;
                Data Source={inFilePath};
                Extended Properties=""Excel 12.0 Xml;HDR=YES""";
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                var cmd = conn.CreateCommand();
                cmd.CommandText = $@"SELECT * FROM [{sheetName}$]";
                using (var dr = cmd.ExecuteReader())
                {
                    var query =
                        (from DbDataRecord row in dr
                            select row).Select(x =>
                        {
                            var data = new Dictionary<string, object>
                            {
                                {dr.GetName(0), x[0]},
                                {dr.GetName(1), x[1]},
                                {dr.GetName(2), x[2]},
                                {dr.GetName(3), x[3]},
                                {dr.GetName(4), x[4]}
                            };
                            return data;
                        });

                    var json = JsonConvert.SerializeObject(query);
                    File.WriteAllText(outFilePath, json);
                }
            }
        }

        private static void UsingExcelDataReader(string inputFile, string outputFile)
        {
            using (var inFile = File.Open(inputFile, FileMode.Open, FileAccess.Read))
            using (var outFile = File.CreateText(outputFile))
            using (var reader = ExcelReaderFactory.CreateReader(inFile,
                new ExcelReaderConfiguration {FallbackEncoding = Encoding.GetEncoding(1252)}))
            using (var writer = new JsonTextWriter(outFile))
            {
                writer.Formatting = Formatting.Indented;
                writer.WriteStartArray();
                //SKIP FIRST ROW, it's TITLES.
                reader.Read();
                do
                {
                    while (reader.Read())
                    {
                        //We don't need empty object
                        var firstName = reader.GetString(0);
                        if (string.IsNullOrEmpty(firstName)) break;

                        writer.WriteStartObject();
                        //Select Columns and values
                        writer.WritePropertyName("FirstName");
                        writer.WriteValue(firstName);

                        writer.WritePropertyName("LastName");
                        writer.WriteValue(reader.GetString(1));

                        writer.WritePropertyName("Gender");
                        writer.WriteValue(reader.GetString(2));

                        writer.WritePropertyName("State");
                        writer.WriteValue(reader.GetString(3));

                        writer.WriteEndObject();
                    }
                } while (reader.NextResult());

                writer.WriteEndArray();
            }
        }
    }
}