// Title: Import JSON array and create a named range with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, imports a JSON array as a table starting at A1 using JsonUtility with ArrayAsTable enabled, defines a named range that covers the imported cells, and saves the file as an XLSX workbook.
// Keywords: Aspose.Cells JSON import C# | JsonUtility ImportData | ArrayAsTable JsonLayoutOptions | Excel named range from JSON | create named range Aspose.Cells | C# import JSON to Excel
// Common Searches: Aspose.Cells import JSON as table C# | How to create a named range after JSON import in Aspose.Cells | JsonUtility ImportData example .NET | Define named range for JSON data in Excel using Aspose | ArrayAsTable option Aspose.Cells
// Developer Intent: Load JSON data into an Excel worksheet and automatically generate a named range that references the imported cells.
// Use Cases: Populate Excel reports from external JSON feeds and reference the data via a named range in formulas or charts. | Use the named range as a source for data‑validation lists or pivot tables after importing JSON records. | Programmatically adjust the named range size when the JSON array length changes at runtime.
// AI Prompts: Write C# code that reads a JSON file, imports it into an Aspose.Cells worksheet as a table, and creates a named range covering the imported area. | Explain the effect of JsonLayoutOptions.ArrayAsTable and how to calculate the correct range dimensions for a named range after JSON import. | Show how to resize an existing named range when the imported JSON array contains a different number of rows.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonImportExample
{
    // Creates a workbook, imports a JSON array as a table starting at A1 using JsonUtility with ArrayAsTable enabled, defines a named range that covers the imported cells, and saves the file as an XLSX workbook.
    class Program
    {
        static void Main()
        {
            // Sample JSON array (each object will become a row)
            string json = @"[
                { ""Name"": ""John"", ""Age"": 30, ""City"": ""New York"" },
                { ""Name"": ""Alice"", ""Age"": 25, ""City"": ""London"" },
                { ""Name"": ""Bob"", ""Age"": 28, ""City"": ""Paris"" }
            ]";

            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Set JSON layout options – treat the array as a table (header + rows)
            JsonLayoutOptions jsonOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true
            };

            // 3. Import JSON data starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(json, sheet.Cells, 0, 0, jsonOptions);

            // 4. Define a named range that covers the imported data.
            //    In this example the JSON has 3 columns (Name, Age, City) and 3 rows (header + 2 data rows).
            //    Adjust the range size if your JSON differs.
            Aspose.Cells.Range dataRange = sheet.Cells.CreateRange("A1:C3");
            dataRange.Name = "MyDataRange";

            // 5. Save the workbook
            workbook.Save("JsonImportedNamedRange.xlsx");
        }
    }
}
