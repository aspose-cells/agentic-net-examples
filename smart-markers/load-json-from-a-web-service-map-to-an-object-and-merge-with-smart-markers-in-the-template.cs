using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

class Program
{
    // Entry point of the application (async to allow awaiting HTTP calls)
    static async Task Main()
    {
        try
        {
            // Path to the Excel template that contains smart markers (e.g., &=$DataSource.Name)
            const string templatePath = "Template.xlsx";

            // Ensure the template file exists; if not, create a minimal workbook as a fallback
            Workbook workbook;
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
                // Optionally add a placeholder smart marker for demonstration
                workbook.Worksheets[0].Cells["A1"].PutValue("&=$DataSource.Name");
                workbook.Save(templatePath);
            }

            // Initialize the WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // URL of the web service that returns JSON data
            const string jsonUrl = "https://example.com/api/data";

            // Retrieve JSON string from the web service (fallback to empty JSON on failure)
            string jsonData = "{}";
            try
            {
                using HttpClient client = new HttpClient();
                jsonData = await client.GetStringAsync(jsonUrl);
            }
            catch (HttpRequestException ex)
            {
                Console.WriteLine($"Warning: Unable to retrieve JSON data ({ex.Message}). Using empty JSON.");
            }

            // Set the JSON string as a data source for smart markers.
            // The name "DataSource" must match the name used in the smart markers inside the template.
            designer.SetJsonDataSource("DataSource", jsonData);

            // Process the smart markers and populate the worksheet with data from the JSON source
            designer.Process();

            // Save the resulting workbook (uses Aspose.Cells save logic)
            const string outputPath = "Result.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook processed and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}