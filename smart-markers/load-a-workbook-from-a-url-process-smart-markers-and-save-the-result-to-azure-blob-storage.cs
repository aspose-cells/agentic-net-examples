// Title: C# example: download an Excel template via HttpClient, apply Aspose.Cells Smart Markers with WorkbookDesigner, and save the processed workbook to a local .xlsx file
// Common Searches: how to use Aspose.Cells WorkbookDesigner to process smart markers after downloading an Excel file with HttpClient in C# | C# code for loading an Excel workbook from a web URL into a MemoryStream and applying smart markers | setting a DataTable as a data source for Aspose.Cells smart markers in .NET | saving a processed Aspose.Cells workbook to a timestamped file on disk | example of downloading an .xlsx template and populating it with data using Aspose.Cells smart markers
// Tags: download Excel workbook with HttpClient C# | Aspose.Cells WorkbookDesigner process smart markers | set DataTable data source for smart markers | save processed workbook as Xlsx file | memory stream workbook loading Aspose.Cells

using System;
using System.Data;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

// The sample downloads an Excel workbook from a specified URL using HttpClient, loads it into a MemoryStream, creates an Aspose.Cells Workbook, prepares a DataTable with sample data, assigns the table to WorkbookDesigner under the name "Data", processes the smart markers, and saves the resulting workbook as a timestamped .xlsx file in the current directory, ensuring the output folder exists and handling any errors.
class Program
{
    // URL of the workbook to load
    private const string WorkbookUrl = "https://example.com/path/to/your/workbook.xlsx";

    static async Task Main()
    {
        try
        {
            // 1. Download the workbook into a memory stream
            byte[] workbookBytes;
            using (var httpClient = new HttpClient())
            {
                workbookBytes = await httpClient.GetByteArrayAsync(WorkbookUrl);
            }

            using var workbookStream = new MemoryStream(workbookBytes);

            // 2. Load the workbook using Aspose.Cells
            var workbook = new Workbook(workbookStream);

            // 3. Prepare the data source for Smart Markers
            DataTable dataTable = GetSmartMarkerData();

            // 4. Process Smart Markers using WorkbookDesigner
            var designer = new WorkbookDesigner(workbook);
            // The name "Data" must match the smart marker reference in the template (e.g., &=Data.ColumnName)
            designer.SetDataSource("Data", dataTable);
            designer.Process();

            // 5. Save the processed workbook to a local file
            string outputPath = Path.Combine(
                Environment.CurrentDirectory,
                $"processed_{DateTime.UtcNow:yyyyMMdd_HHmmss}.xlsx");

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook processed and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Placeholder method to obtain data for Smart Markers.
    // Replace this with actual data retrieval (e.g., from a database, API, etc.).
    private static DataTable GetSmartMarkerData()
    {
        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Quantity", typeof(int));
        table.Columns.Add("Price", typeof(decimal));

        // Sample rows
        table.Rows.Add("Apple", 10, 0.5m);
        table.Rows.Add("Banana", 20, 0.3m);
        table.Rows.Add("Cherry", 15, 1.2m);

        return table;
    }
}
