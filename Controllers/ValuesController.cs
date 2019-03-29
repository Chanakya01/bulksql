using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web;
using ExcelDataReader;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Server.Kestrel.Core.Internal.Infrastructure;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;

namespace infoapi.Controllers {
    [Route ("api/[controller]")]
    [ApiController]
    public class ValuesController : ControllerBase {
        public string connectionstr = "Server=tcp:bulktable.database.windows.net,1433;Initial Catalog=bulktable;Persist Security Info=False;User ID=chanakya;Password=jahnavi@01;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;";
        // GET api/values
        [HttpGet]
        public ActionResult<IEnumerable<string>> Get () {
            using (SqlConnection conn = new SqlConnection ()) {
                conn.ConnectionString = connectionstr;
                conn.Open ();
                List<string> tables = new List<string> ();
                DataTable dt = conn.GetSchema ("Tables");
                foreach (DataRow row in dt.Rows) {
                    string tablename = (string) row[2];
                    if (tablename != "master" && tablename !="database_firewall_rules")
                        tables.Add (tablename);
                }
                return tables;
            }
        }

        [HttpPost]
        [Route ("Postcol")]
        public ActionResult<IEnumerable<string>> Postcol (List<string> id) {
            // getmaster ();
            return getcolumnmap (id).ToArray ();
        }

        public List<string> getcolumnmap (List<string> tableid) {

            List<string> listacolumnas = new List<string> ();
            foreach (var li in tableid) {
                using (SqlConnection connection = new SqlConnection (connectionstr))
                using (SqlCommand command = connection.CreateCommand ()) {
                    command.CommandText = string.Format ("select c.name from sys.columns c inner join sys.tables t on t.object_id = c.object_id and t.name = '{0}' and t.type = 'U'", li);
                    connection.Open ();
                    using (var reader = command.ExecuteReader ()) {
                        while (reader.Read ()) {
                            DataTable table = getmaster ("ColumnName", reader.GetString (0));
                            listacolumnas.Add (string.Format("{0}:  {1}",table.Rows[0]["Name"].ToString (),li));
                        }
                    }
                }

            }
            return listacolumnas;
        }
        //POST api/values
        [HttpPost]
        public ActionResult Post ([FromBody] JArray value, [FromQuery] string tableName) {
            using (ExcelPackage excel = new ExcelPackage ()) {
                excel.Workbook.Worksheets.Add ("Worksheet1");

                var headerRow = new List<string[]> ();
                var editer = new List<string> ();
                var items = value.Select (jv => (string) jv);
                foreach(var i in items)
                {
                    editer.Add(i.Split(':')[0]);
                }
                headerRow.Add (editer.ToArray());
                string headerRange = "A1:" + Char.ConvertFromUtf32 (headerRow[0].Length + 64) + "1";

                var worksheet = excel.Workbook.Worksheets["Worksheet1"];

                worksheet.Cells[headerRange].LoadFromArrays (headerRow);
                worksheet.Cells.AutoFitColumns();
                MemoryStream stream = new MemoryStream ();
                excel.SaveAs (stream);
                stream.Seek (0, SeekOrigin.Begin);

                byte[] bytes = new byte[stream.Length];
                stream.Read (bytes, 0, bytes.Length);

                var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                var donwloadFile = string.Format ("{0}.xlsx", tableName);

                return File (bytes, contentType, donwloadFile);
            }
        }

        [HttpPost]
        [Route ("Postfile")]
        public ActionResult PostFile () {
            string tableName;
            var file = Request.Form.Files[0];

            System.Text.Encoding.RegisterProvider (System.Text.CodePagesEncodingProvider.Instance);
            using (Stream stream = file.OpenReadStream ()) {

                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader (stream);
                var result = excelReader.AsDataSet (new ExcelDataSetConfiguration () {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration () {
                        UseHeaderRow = true
                    }
                });
                Dictionary<string, List<string>> dic = new Dictionary<string, List<string>> ();
                DataTable firstTable = result.Tables[0];
                foreach (DataColumn col in firstTable.Columns) {
                    DataTable table = getmaster ("Name", col.ColumnName);

                    tableName = table.Rows[0]["TableName"].ToString ();
                    if (!dic.ContainsKey (tableName)) {
                        List<string> list = new List<string> ();
                        list.Add (col.ColumnName);
                        dic.Add (tableName, list);
                    } else {
                        List<string> list1 = new List<string> ();
                        dic.TryGetValue (tableName, out list1);
                        list1.Add (col.ColumnName);
                    }
                }
                foreach (var item in dic) {
                    List<string> de = item.Value;
                    var bulkCopy = new SqlBulkCopy (connectionstr);
                    bulkCopy.DestinationTableName = item.Key;
                    foreach (string col in de) {
                        DataTable table = getmaster ("Name", col);
                        string last = table.Rows[0]["ColumnName"].ToString ();
                        SqlBulkCopyColumnMapping mapID = new SqlBulkCopyColumnMapping (col, last);
                        bulkCopy.ColumnMappings.Add (mapID);
                    }
                    bulkCopy.WriteToServer (firstTable);
                }
                return Ok ();
            }
        }
        public DataTable getmaster (string key, string colname) {
            using (SqlConnection conn = new SqlConnection (connectionstr)) {
                conn.Open ();
                string query = string.Format ("SELECT * FROM bulktable. dbo.master where {0} like '{1}'", key, colname);
                var table = new DataTable ();
                using (var dara = new SqlDataAdapter (query, connectionstr)) {
                    dara.Fill (table);
                }
                return table;
            }
        }
        // PUT api/values/5
        [HttpPut ("{id}")]
        public void Put (int id, [FromBody] string value) { }

        // DELETE api/values/5
        [HttpDelete ("{id}")]
        public void Delete (int id) { }
    }
}