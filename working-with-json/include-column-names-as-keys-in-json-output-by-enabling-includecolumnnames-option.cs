// Title: C# – Export Excel range to JSON with header row as property names using Aspose.Cells for .NET
// Description: A concise C# example that creates a workbook, adds a header row (Name, Age, Country) and two data rows, defines the range A1:C3, and calls Aspose.Cells JsonUtility.ExportRangeToJson with ExportRangeToJsonOptions (HasHeaderRow = true, Indent = " ") to generate a pretty‑printed JSON string where the first row supplies the JSON keys.
// Keywords: Aspose.Cells export range to JSON | ExportRangeToJsonOptions HasHeaderRow | C# Excel to JSON conversion | pretty printed JSON Aspose.Cells | JsonUtility ExportRangeToJson | .NET Excel JSON example | header row as JSON keys | range A1:C3 JSON export
// Common Searches: How to export an Excel range to JSON with column headers using Aspose.Cells C# | Aspose.Cells ExportRangeToJsonOptions HasHeaderRow true example | C# pretty‑print JSON from worksheet range Aspose.Cells | Export multiple Excel ranges to separate JSON objects Aspose.Cells | GitHub sample for Excel to JSON with Aspose.Cells
// Developer Intent: Generate a formatted JSON string from a selected Excel range, using the first row as the JSON property names.
// Use Cases: Create API payloads directly from small lookup tables stored in Excel. | Convert configuration data maintained in spreadsheets into JSON files for .NET services. | Produce test data sets in JSON format without manual copy‑paste.
// AI Prompts: Write C# code that uses Aspose.Cells to export a worksheet range to indented JSON, treating the first row as keys. | Show how to modify ExportRangeToJsonOptions to output compact (non‑indented) JSON from Excel. | Explain how to export several non‑contiguous ranges to separate JSON objects with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

// A concise C# example that creates a workbook, adds a header row (Name, Age, Country) and two data rows, defines the range A1:C3, and calls Aspose.Cells JsonUtility.ExportRangeToJson with ExportRangeToJsonOptions (HasHeaderRow = true, Indent = " ") to generate a pretty‑printed JSON string where the first row supplies the JSON keys.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Add header row (column names)
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Age");
            cells["C1"].PutValue("Country");

            // Add sample data rows
            cells["A2"].PutValue("John");
            cells["B2"].PutValue(30);
            cells["C2"].PutValue("USA");

            cells["A3"].PutValue("Alice");
            cells["B3"].PutValue(25);
            cells["C3"].PutValue("UK");

            // Define the range that includes the header and data
            AsposeRange range = cells.CreateRange("A1:C3");

            // Configure JSON export options to treat the first row as header (column names as keys)
            ExportRangeToJsonOptions exportOptions = new ExportRangeToJsonOptions
            {
                HasHeaderRow = true,   // Use column names as JSON keys
                Indent = "  "          // Pretty‑print with indentation
            };

            // Export the range to a JSON string
            string jsonOutput = JsonUtility.ExportRangeToJson(range, exportOptions);

            // Display the resulting JSON
            Console.WriteLine(jsonOutput);
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
