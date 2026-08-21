// Title: Aspose.Cells .NET: Access a Worksheet by Its Name and Obtain a Reference (C#)
// Description: Demonstrates how to create a workbook, add a sheet named "DataSheet", retrieve that sheet using the Workbook.Worksheets["DataSheet"] indexer, write a value to cell A1, and save the file as AccessWorksheetByName.xlsx.
// Keywords: Aspose.Cells | C# | access worksheet by name | retrieve worksheet reference | Workbook.Worksheets indexer | get sheet by string key | Aspose.Cells .NET example | worksheet name lookup | C# spreadsheet library
// Common Searches: Aspose.Cells get worksheet by name C# | How to retrieve a sheet using its name in Aspose.Cells | Workbook.Worksheets["SheetName"] example | C# Aspose.Cells access specific worksheet | Aspose.Cells reference worksheet after adding
// Developer Intent: Fetch a Worksheet object from a Workbook by specifying the sheet's name.
// Use Cases: After adding a custom‑named sheet, locate it later to populate data or apply formatting. | Use the named worksheet reference to write values, formulas, or styles to particular cells. | Check for a sheet's existence by name before performing operations such as data import or validation. | Reuse the worksheet reference across multiple methods to keep code clean and avoid repeated lookups.
// AI Prompts: Write C# code that verifies a worksheet named "Report" exists before accessing it with Aspose.Cells. | Show how to loop through all worksheets in a workbook and apply formatting only to the sheet called "Summary" using Aspose.Cells for .NET. | Explain best practices for handling exceptions when retrieving a worksheet by name in Aspose.Cells. | Generate a method that returns a Worksheet object given a workbook and a sheet name, creating the sheet if it does not already exist.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a sheet named "DataSheet", retrieve that sheet using the Workbook.Worksheets["DataSheet"] indexer, write a value to cell A1, and save the file as AccessWorksheetByName.xlsx.
    public class AccessWorksheetByNameExample
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a worksheet with a specific name
                Worksheet newSheet = workbook.Worksheets.Add("DataSheet");

                // Access the worksheet by its name
                Worksheet accessedSheet = workbook.Worksheets["DataSheet"];

                // Write a message to cell A1
                accessedSheet.Cells["A1"].PutValue("Worksheet accessed successfully!");

                // Save the workbook
                string outputPath = "AccessWorksheetByName.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
