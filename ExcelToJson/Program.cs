using System.Collections.Generic;
using System.Data.Common;
using System.Data.OleDb;
using System.IO;
using System.Linq;
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
    }
}