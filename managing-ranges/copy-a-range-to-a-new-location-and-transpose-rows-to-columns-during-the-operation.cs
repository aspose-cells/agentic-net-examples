// Title: Copy and Transpose a Range in Excel with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to copy a source range to a new location and transpose rows into columns using Aspose.Cells. The example creates a workbook, fills a horizontal range, configures PasteOptions with Transpose=true, copies the data to a vertical range, and saves the file.
// Keywords: Aspose.Cells | C# copy range | transpose rows to columns | PasteOptions Transpose | Excel range copy .NET | Aspose.Cells example | range transposition | copy range to new location | Aspose.Cells PasteOptions | Excel automation C#
// Common Searches: Aspose.Cells copy range and transpose | How to transpose data while copying in Aspose.Cells C# | PasteOptions Transpose property example | Copy horizontal range to vertical column Aspose.Cells | C# Aspose.Cells range copy with formatting
// Developer Intent: Copy a source range to another worksheet location and transpose its rows into columns using Aspose.Cells.
// Use Cases: Convert a row of month headers into a vertical list for reporting dashboards. | Reorient a data table so that rows become columns before generating a chart. | Duplicate a formatted block while flipping its orientation to fit a different layout.
// AI Prompts: Generate C# code that copies a range and transposes it with Aspose.Cells, preserving formatting and formulas. | Explain how PasteOptions.Transpose works with different PasteType values in Aspose.Cells. | Show an example of copying multiple rows and transposing them into columns, including robust error handling.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsTransposeExample
{
    // Demonstrates how to copy a source range to a new location and transpose rows into columns using Aspose.Cells. The example creates a workbook, fills a horizontal range, configures PasteOptions with Transpose=true, copies the data to a vertical range, and saves the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Fill sample data in a horizontal range (A1:D1)
                cells["A1"].PutValue("Jan");
                cells["B1"].PutValue("Feb");
                cells["C1"].PutValue("Mar");
                cells["D1"].PutValue("Apr");

                // Define the source range (A1:D1)
                AsposeRange sourceRange = cells.CreateRange("A1:D1");

                // Create PasteOptions and enable transposition
                PasteOptions pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.All, // copy everything (values, formats, etc.)
                    Transpose = true           // transpose rows ↔ columns
                };

                // Define the destination range where the transposed data will be placed (A2:A5)
                AsposeRange destinationRange = cells.CreateRange("A2:A5");

                // Copy the source range to the destination range with transposition
                destinationRange.Copy(sourceRange, pasteOptions);

                // Determine output file path
                string outputFile = "TransposedCopy.xlsx";

                // Save the workbook to a file
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputFile)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
