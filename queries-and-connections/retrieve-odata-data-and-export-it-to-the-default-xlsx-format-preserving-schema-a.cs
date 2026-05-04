using System;
using System.Data;
using System.Net.Http;
using System.Threading.Tasks;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace ODataToExcelExample
{
    class Program
    {
        static async Task Main(string[] args)
        {
            string odataUrl = "https://services.odata.org/V4/Northwind/Northwind.svc/Products";

            DataTable dataTable = await GetODataDataTableAsync(odataUrl);

            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = true,
                InsertRows = true,
                ConvertNumericData = true
            };
            sheet.Cells.ImportData(dataTable, 0, 0, importOptions);

            workbook.Save("ODataExport.xlsx");
            Console.WriteLine("OData data exported successfully to ODataExport.xlsx");
        }

        private static async Task<DataTable> GetODataDataTableAsync(string requestUrl)
        {
            using HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(
                new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

            HttpResponseMessage response = await client.GetAsync(requestUrl);
            response.EnsureSuccessStatusCode();

            string rawJson = await response.Content.ReadAsStringAsync();
            using JsonDocument doc = JsonDocument.Parse(rawJson);
            JsonElement root = doc.RootElement;

            if (!root.TryGetProperty("value", out JsonElement valueElement) || valueElement.ValueKind != JsonValueKind.Array)
                throw new InvalidOperationException("Invalid OData response format.");

            return BuildDataTableFromJsonArray(valueElement);
        }

        private static DataTable BuildDataTableFromJsonArray(JsonElement arrayElement)
        {
            DataTable table = new DataTable();

            bool columnsDefined = false;
            foreach (JsonElement item in arrayElement.EnumerateArray())
            {
                if (!columnsDefined)
                {
                    foreach (JsonProperty prop in item.EnumerateObject())
                    {
                        Type columnType = GetClrTypeFromJsonValue(prop.Value);
                        table.Columns.Add(prop.Name, columnType);
                    }
                    columnsDefined = true;
                }

                DataRow row = table.NewRow();
                foreach (JsonProperty prop in item.EnumerateObject())
                {
                    object value = GetClrValueFromJsonValue(prop.Value);
                    row[prop.Name] = value ?? DBNull.Value;
                }
                table.Rows.Add(row);
            }

            return table;
        }

        private static Type GetClrTypeFromJsonValue(JsonElement element)
        {
            return element.ValueKind switch
            {
                JsonValueKind.Number => element.TryGetInt64(out _) ? typeof(long) :
                                         element.TryGetDouble(out _) ? typeof(double) :
                                         typeof(decimal),
                JsonValueKind.String => typeof(string),
                JsonValueKind.True or JsonValueKind.False => typeof(bool),
                JsonValueKind.Null => typeof(string),
                _ => typeof(string)
            };
        }

        private static object GetClrValueFromJsonValue(JsonElement element)
        {
            return element.ValueKind switch
            {
                JsonValueKind.Number => element.TryGetInt64(out long l) ? (object)l :
                                         element.TryGetDouble(out double d) ? d :
                                         element.GetDecimal(),
                JsonValueKind.String => element.GetString(),
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                JsonValueKind.Null => null,
                _ => element.GetRawText()
            };
        }
    }
}