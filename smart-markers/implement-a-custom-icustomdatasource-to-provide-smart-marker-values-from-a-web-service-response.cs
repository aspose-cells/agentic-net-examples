// Title: C# Custom ICellsDataTable for Aspose.Cells Smart Markers – Load JSON Data from a Web Service
// Description: Demonstrates a C# ICellsDataTable implementation that fetches product data via HttpClient, deserializes a JSON array, and supplies Name and Price columns to WorkbookDesigner smart markers, with fallback sample data and Excel file output.
// Keywords: Aspose.Cells | ICellsDataTable | smart markers | C# | web service | JSON API | REST | WorkbookDesigner | custom data source | Excel export | product catalog | HttpClient
// Common Searches: Aspose.Cells custom ICellsDataTable example | populate Excel smart markers from REST API C# | load JSON data into Aspose.Cells workbook | WorkbookDesigner set data source from web service | fallback data for Aspose.Cells smart markers
// Developer Intent: Create a reusable ICellsDataTable that retrieves product information from a JSON web service and feeds it to smart markers in an Aspose.Cells workbook.
// Use Cases: Generate a product catalog Excel file by pulling the latest items from a live REST endpoint. | Automate sales or inventory reports where product names and prices are inserted via smart markers. | Ensure report generation continues when the API is unavailable by providing built‑in fallback data. | Integrate Aspose.Cells into ASP.NET or console applications that need dynamic Excel output from external services.
// AI Prompts: Write a C# class that implements ICellsDataTable, reads a JSON array from a URL, and exposes column names for smart markers. | Show how to bind the custom data source to WorkbookDesigner, place smart markers like &=$Products.Name, and generate an Excel file. | Explain error handling strategies and how to supply fallback sample data when the web service call fails.

using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsCustomDataSourceDemo
{
    // Simple POCO representing a product returned by the web service
    // Demonstrates a C# ICellsDataTable implementation that fetches product data via HttpClient, deserializes a JSON array, and supplies Name and Price columns to WorkbookDesigner smart markers, with fallback sample data and Excel file output.
    public class Product
    {
        public string Name { get; set; }
        public double Price { get; set; }
    }

    // Custom data source implementing ICellsDataTable.
    // It fetches data from a web service (JSON array of Product objects) and
    // provides the required members for smart marker processing.
    public class WebServiceDataSource : ICellsDataTable
    {
        private readonly List<Product> _products;
        private int _currentRow = -1;

        // Constructor accepts the endpoint URL, performs the HTTP GET request,
        // and deserializes the JSON response into a list of Product objects.
        // If the request fails, fallback sample data is used.
        public WebServiceDataSource(string requestUrl)
        {
            List<Product> products = null;

            try
            {
                using var httpClient = new HttpClient();
                var response = httpClient.GetAsync(requestUrl).Result;
                response.EnsureSuccessStatusCode();
                string json = response.Content.ReadAsStringAsync().Result;

                // Expecting JSON like: [{ "Name": "Item1", "Price": 10.5 }, ...]
                products = JsonSerializer.Deserialize<List<Product>>(json,
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
            }
            catch (Exception)
            {
                // Fallback to sample data when the web service is unavailable
                products = new List<Product>
                {
                    new Product { Name = "Sample Product 1", Price = 9.99 },
                    new Product { Name = "Sample Product 2", Price = 19.99 }
                };
            }

            _products = products ?? new List<Product>();
        }

        // Indexer for row/column access (column 0 = Name, column 1 = Price)
        public object this[int rowIndex, int columnIndex]
        {
            get
            {
                var product = _products[rowIndex];
                return columnIndex == 0 ? product.Name : (object)product.Price;
            }
        }

        // Indexer for row access – returns the whole object for the current row.
        public object this[int rowIndex] => _products[rowIndex];

        // Indexer for column name access (used by smart markers like $Products.Name)
        public object this[string columnName]
        {
            get
            {
                // _currentRow is set by Next()
                var product = _products[_currentRow];
                return columnName switch
                {
                    "Name" => product.Name,
                    "Price" => product.Price,
                    _ => null
                };
            }
        }

        // Total number of rows
        public int RowCount => _products.Count;

        // Number of columns (Name, Price)
        public int ColumnCount => 2;

        // Alias for RowCount
        public int Count => _products.Count;

        // Column names exposed to the designer
        public string[] Columns => new[] { "Name", "Price" };

        // Reset enumeration to before the first row
        public void BeforeFirst()
        {
            _currentRow = -1;
        }

        // Move to the next row; returns false when no more rows are available
        public bool Next()
        {
            _currentRow++;
            return _currentRow < _products.Count;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (template)
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Place smart markers in the first two rows (header + data row)
                sheet.Cells["A1"].PutValue("&=$Products.Name");
                sheet.Cells["B1"].PutValue("&=$Products.Price");

                // Initialize the designer with the workbook
                var designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Set the custom data source. Replace the URL with a real endpoint.
                // For demonstration, you can host a simple JSON file locally.
                string serviceUrl = "https://example.com/api/products";
                designer.SetDataSource("Products", new WebServiceDataSource(serviceUrl));

                // Process smart markers – this will populate the cells with data from the web service
                designer.Process();

                // Determine output path and ensure the directory exists
                string outputPath = "SmartMarkerFromWebService.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the resulting workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
