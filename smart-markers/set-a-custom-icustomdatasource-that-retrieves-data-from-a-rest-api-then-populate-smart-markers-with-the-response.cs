// Title: C# Custom ICellsDataTable that pulls JSON from a REST API to fill Aspose.Cells smart markers
// Description: Demonstrates how to implement a custom ICellsDataTable, fetch a list of Person objects from a REST endpoint (with fallback sample data), bind the data source to WorkbookDesigner, process smart markers (&=$Person.Name, &=$Person.Age, &=$Person.City) and generate an Excel report.
// Keywords: Aspose.Cells custom data source | ICellsDataTable C# example | smart markers REST API | populate Excel from JSON | WorkbookDesigner SetDataSource | C# HTTP client JSON deserialization | Excel report generation Aspose
// Common Searches: how to bind REST API data to Aspose.Cells smart markers | implement ICellsDataTable for JSON in .NET | Aspose.Cells custom data source tutorial | C# generate Excel from web service | smart markers with custom data source Aspose
// Developer Intent: Retrieve JSON data from a web service, expose it through a custom ICellsDataTable, and automatically fill smart markers in an Excel template.
// Use Cases: Create an employee directory by consuming an HR API and exporting to a formatted workbook. | Generate sales or inventory dashboards that pull live JSON feeds into pre‑designed Excel templates. | Build automated reporting tools that merge external REST data with Aspose.Cells smart markers for scheduled Excel outputs.
// AI Prompts: Write a C# method that calls a REST endpoint and returns a List<T> suitable for a custom ICellsDataTable. | Explain how to map JSON property names to column names required by Aspose.Cells smart markers when implementing ICellsDataTable. | Provide robust error‑handling patterns for WorkbookDesigner.Process when the external API is unavailable.

using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsCustomDataSourceDemo
{
    // Model representing the JSON objects returned by the REST API
    // Demonstrates how to implement a custom ICellsDataTable, fetch a list of Person objects from a REST endpoint (with fallback sample data), bind the data source to WorkbookDesigner, process smart markers (&=$Person.Name, &=$Person.Age, &=$Person.City) and generate an Excel report.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string City { get; set; }
    }

    // Custom data source implementing ICellsDataTable.
    // It fetches data from a REST API, stores it locally, and provides
    // the required members for Aspose.Cells smart markers.
    public class PersonDataSource : ICellsDataTable
    {
        private readonly List<Person> _persons;
        private int _currentRow = -1;

        public PersonDataSource(List<Person> persons)
        {
            _persons = persons ?? new List<Person>();
        }

        // Indexer for row/column access (used by smart markers)
        public object this[int rowIndex, int columnIndex]
        {
            get
            {
                var person = _persons[rowIndex];
                return columnIndex switch
                {
                    0 => person.Name,
                    1 => person.Age,
                    2 => person.City,
                    _ => null
                };
            }
        }

        // Indexer for row access (required by the interface)
        public object this[int rowIndex] => _persons[rowIndex];

        // Indexer for column name access (used by smart markers)
        public object this[string columnName]
        {
            get
            {
                var person = _persons[_currentRow];
                return columnName switch
                {
                    "Name" => person.Name,
                    "Age" => person.Age,
                    "City" => person.City,
                    _ => null
                };
            }
        }

        public int RowCount => _persons.Count;
        public int ColumnCount => 3; // Name, Age, City
        public int Count => _persons.Count;
        public string[] Columns => new[] { "Name", "Age", "City" };

        public void BeforeFirst()
        {
            _currentRow = -1;
        }

        public bool Next()
        {
            _currentRow++;
            return _currentRow < _persons.Count;
        }
    }

    class Program
    {
        // Asynchronous method to fetch data from a REST endpoint.
        // Returns sample data if the request fails.
        private static async Task<List<Person>> FetchPersonsFromApiAsync()
        {
            const string apiUrl = "https://example.com/api/persons"; // replace with real URL

            try
            {
                using HttpClient client = new HttpClient();
                HttpResponseMessage response = await client.GetAsync(apiUrl);
                response.EnsureSuccessStatusCode();

                string json = await response.Content.ReadAsStringAsync();
                return JsonSerializer.Deserialize<List<Person>>(json, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                }) ?? new List<Person>();
            }
            catch (Exception ex)
            {
                // Log the error (could be replaced with a proper logging framework)
                Console.WriteLine($"Warning: Unable to fetch data from API. {ex.Message}");
                // Return fallback sample data
                return new List<Person>
                {
                    new Person { Name = "John Doe", Age = 30, City = "New York" },
                    new Person { Name = "Jane Smith", Age = 25, City = "London" },
                    new Person { Name = "Carlos Ruiz", Age = 40, City = "Madrid" }
                };
            }
        }

        static async Task Main()
        {
            try
            {
                // 1. Create a new workbook and add smart markers.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Header row
                sheet.Cells["A1"].PutValue("Name");
                sheet.Cells["B1"].PutValue("Age");
                sheet.Cells["C1"].PutValue("City");

                // Smart marker rows (the designer will repeat these rows for each data item)
                sheet.Cells["A2"].PutValue("&=$Person.Name");
                sheet.Cells["B2"].PutValue("&=$Person.Age");
                sheet.Cells["C2"].PutValue("&=$Person.City");

                // 2. Retrieve data from the REST API (or fallback data).
                List<Person> persons = await FetchPersonsFromApiAsync();

                // 3. Create the custom data source and bind it to the designer.
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("Person", new PersonDataSource(persons));

                // 4. Process smart markers and populate the worksheet.
                designer.Process();

                // 5. Save the result.
                const string outputPath = "PersonsReport.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Report generated successfully: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
