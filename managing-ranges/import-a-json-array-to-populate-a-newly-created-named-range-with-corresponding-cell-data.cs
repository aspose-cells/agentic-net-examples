// Title: C# – Import a JSON array into Excel and create a named range with Aspose.Cells
// Description: Demonstrates how to use Aspose.Cells for .NET to import a JSON array as a table (with headers) starting at A1, calculate the occupied area, create a named range that covers the imported cells, assign the name "MyJsonData", and save the workbook as an XLSX file.
// Keywords: Aspose.Cells C# import JSON | JsonUtility ImportData example | create named range Aspose.Cells | JSON to Excel table .NET | AsposeRange assign name | C# Excel named range from JSON | Aspose.Cells workbook save | Excel automation JSON data | Aspose.Cells range creation | C# Excel API JSON import
// Common Searches: How to import JSON array into Excel using Aspose.Cells C# | Create a named range after importing JSON with Aspose.Cells | Aspose.Cells JsonUtility ImportData named range example | C# code to convert JSON to Excel table and name the range | Aspose.Cells create range from dynamic data
// Developer Intent: Generate a named range that encapsulates JSON‑derived cells in an Excel workbook.
// Use Cases: Transform API‑provided JSON into a structured Excel table for reporting. | Expose imported JSON data through a named range for formulas, charts, or pivot tables. | Automate workbook creation where downstream tools reference a stable named range.
// AI Prompts: Write C# code with Aspose.Cells to import a JSON array as a table and assign a named range to the imported cells. | Show how to determine the size of JSON‑imported data and create a range that covers the entire block. | Provide an example that saves an XLSX file after creating a named range from JSON data using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to use Aspose.Cells for .NET to import a JSON array as a table (with headers) starting at A1, calculate the occupied area, create a named range that covers the imported cells, assign the name "MyJsonData", and save the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // JSON array to be imported
            string json = @"[
                { ""Name"": ""John"", ""Age"": 30 },
                { ""Name"": ""Alice"", ""Age"": 25 }
            ]";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Configure JSON layout to treat the array as a table (adds header row)
            JsonLayoutOptions layoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true
            };

            // Import JSON data starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(json, cells, 0, 0, layoutOptions);

            // Determine the size of the imported data (including header row)
            int lastRow = cells.MaxDataRow;          // zero‑based index of the last row with data
            int lastColumn = cells.MaxDataColumn;    // zero‑based index of the last column with data

            // Create a range that covers the imported data
            // Parameters: firstRow, firstColumn, totalRows, totalColumns
            AsposeRange importedRange = cells.CreateRange(0, 0, lastRow + 1, lastColumn + 1);

            // Assign a name to the range
            importedRange.Name = "MyJsonData";

            // Save the workbook
            string outputPath = "JsonImported.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
