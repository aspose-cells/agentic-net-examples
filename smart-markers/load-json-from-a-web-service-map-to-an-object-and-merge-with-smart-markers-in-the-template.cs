// Title: Retrieve JSON from a REST API, map to C# objects, and populate Excel smart markers with Aspose.Cells
// Common Searches: how to fill Aspose.Cells smart markers using JSON data from a REST endpoint in C# | C# example for converting a JSON array to a DataSet for WorkbookDesigner | populate an Excel template with product data retrieved via HttpClient and Aspose.Cells | using Aspose.Cells WorkbookDesigner to merge web service data into an .xlsx file | error handling for missing template file and HTTP request when processing smart markers
// Tags: Aspose.Cells WorkbookDesigner JSON integration | Excel smart markers population from web API | C# JSON to DataSet conversion for Aspose | REST API data merge into Excel template | smart marker processing with HttpClient

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // Data model matching the JSON structure
    // The sample loads an Excel workbook containing smart markers, fetches product data as JSON from a REST API, deserializes it into C# objects, converts the collection to a DataSet, assigns the DataSet to WorkbookDesigner, processes the smart markers, and saves the populated file while handling file, HTTP, and JSON errors.
    public class Product
    {
        public string Name { get; set; } = default!;
        public double Price { get; set; }
        public int Quantity { get; set; }
    }

    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                const string templatePath = "Template.xlsx";
                const string resultPath = "Result.xlsx";

                // Ensure the template file exists before loading
                if (!File.Exists(templatePath))
                {
                    throw new FileNotFoundException($"The template file '{templatePath}' was not found.");
                }

                // Load the Excel template that contains smart markers
                var workbook = new Workbook(templatePath);

                // Retrieve JSON data from a web service
                using var httpClient = new HttpClient();
                string json = await httpClient.GetStringAsync("https://example.com/api/products");

                // Map JSON to a list of Product objects
                List<Product> products = JsonSerializer.Deserialize<List<Product>>(json) ?? new List<Product>();

                // Convert the list to a DataSet (required by WorkbookDesigner in this version)
                var dataTable = new DataTable("Products");
                dataTable.Columns.Add("Name", typeof(string));
                dataTable.Columns.Add("Price", typeof(double));
                dataTable.Columns.Add("Quantity", typeof(int));

                foreach (var p in products)
                {
                    dataTable.Rows.Add(p.Name, p.Price, p.Quantity);
                }

                var dataSet = new DataSet();
                dataSet.Tables.Add(dataTable);

                // Process smart markers using WorkbookDesigner
                var designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dataSet);
                designer.Process();

                // Save the populated workbook
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
            }
            catch (FileNotFoundException ex)
            {
                Console.Error.WriteLine($"File error: {ex.Message}");
            }
            catch (HttpRequestException ex)
            {
                Console.Error.WriteLine($"HTTP request error: {ex.Message}");
            }
            catch (JsonException ex)
            {
                Console.Error.WriteLine($"JSON deserialization error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
