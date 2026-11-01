// Title: How to download JSON data with HttpClient and write it into a worksheet cell using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells example for inserting JSON string into a specific cell | download JSON from a web API and place it in Excel using Aspose.Cells | Aspose.Cells .NET code to save workbook after adding external JSON content | using HttpClient to fetch JSON and populate an Aspose.Cells worksheet | write raw JSON response to cell A1 in an Aspose.Cells workbook
// Tags: download JSON with HttpClient C# | write JSON string to worksheet cell Aspose.Cells | save workbook as XLSX Aspose.Cells | Aspose.Cells insert external data into worksheet | C# fetch API data and populate Excel

using System;
using System.Net.Http;
using Aspose.Cells;

// The sample creates a new workbook, renames the first worksheet to "DataSheet", retrieves JSON data from a specified URL using HttpClient, inserts the JSON text into cell A1 when available, and saves the workbook as WebQueryJsonExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook with a default worksheet
            Workbook workbook = new Workbook();

            // Get reference to the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // URL that returns JSON data
            string jsonUrl = "https://api.example.com/data.json";

            // Download JSON data safely
            string jsonData = string.Empty;
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    // Synchronous call to keep Main simple
                    jsonData = client.GetStringAsync(jsonUrl).GetAwaiter().GetResult();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to download JSON data: {ex.Message}");
                // Continue with empty data or exit as needed
            }

            // Place the raw JSON into cell A1 (or process further as required)
            if (!string.IsNullOrEmpty(jsonData))
            {
                sheet.Cells["A1"].PutValue(jsonData);
            }

            // Save the workbook
            workbook.Save("WebQueryJsonExample.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
