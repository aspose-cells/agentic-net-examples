using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsCustomDataSourceDemo
{
    // Model representing the JSON data returned by the REST API
    public class User
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Email { get; set; }
    }

    // Custom data source implementing ICellsDataTable
    public class ApiDataSource : ICellsDataTable
    {
        private readonly List<User> _users;
        private int _currentRow = -1;

        // Column names in the order they will be accessed by the smart markers
        private static readonly string[] _columns = { "Id", "Name", "Email" };

        public ApiDataSource(List<User> users)
        {
            _users = users ?? new List<User>();
        }

        // Indexer for row/column access by index
        public object this[int rowIndex, int columnIndex]
        {
            get
            {
                var user = _users[rowIndex];
                return columnIndex switch
                {
                    0 => user.Id,
                    1 => user.Name,
                    2 => user.Email,
                    _ => null
                };
            }
        }

        // Indexer for row access (returns the whole object)
        public object this[int rowIndex] => _users[rowIndex];

        // Indexer for column access by name (uses the current row)
        public object this[string columnName]
        {
            get
            {
                var user = _users[_currentRow];
                return columnName switch
                {
                    "Id" => user.Id,
                    "Name" => user.Name,
                    "Email" => user.Email,
                    _ => null
                };
            }
        }

        public int RowCount => _users.Count;
        public int ColumnCount => _columns.Length;
        public int Count => _users.Count;
        public string[] Columns => _columns;

        public void BeforeFirst()
        {
            _currentRow = -1;
        }

        public bool Next()
        {
            _currentRow++;
            return _currentRow < _users.Count;
        }
    }

    class Program
    {
        static void Main()
        {
            // Step 1: Retrieve data from a REST API
            var users = FetchUsersFromApi();

            // Step 2: Create a new workbook and set up smart markers
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Email");

            // Smart marker row (will be repeated for each record)
            sheet.Cells["A2"].PutValue("&=$ApiData.Id");
            sheet.Cells["B2"].PutValue("&=$ApiData.Name");
            sheet.Cells["C2"].PutValue("&=$ApiData.Email");

            // Step 3: Bind the custom data source to the workbook designer
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;
            designer.SetDataSource("ApiData", new ApiDataSource(users));

            // Step 4: Process smart markers and save the result
            designer.Process();
            workbook.Save("ApiDataOutput.xlsx");
        }

        // Helper method to call the REST API and deserialize the JSON response
        private static List<User> FetchUsersFromApi()
        {
            const string apiUrl = "https://jsonplaceholder.typicode.com/users";

            using HttpClient client = new HttpClient();
            string json = client.GetStringAsync(apiUrl).Result;

            // Deserialize JSON array into a list of User objects
            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            return JsonSerializer.Deserialize<List<User>>(json, options) ?? new List<User>();
        }
    }
}