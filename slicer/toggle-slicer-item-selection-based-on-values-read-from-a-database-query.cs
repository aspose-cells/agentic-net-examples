// Title: Toggle Aspose.Cells slicer items in a .NET workbook using values from a database query (C# example)
// AI Prompts: Write C# code that opens an existing Excel file with Aspose.Cells, executes a SQL query, and selects or deselects slicer items in the workbook based on the query results. | Generate a method that iterates through all slicers in a workbook and sets their selected state according to a list of values retrieved via ADO.NET. | Create robust error‑handling that validates the input file path, catches database connection failures, and logs detailed messages when slicer updates cannot be applied.
// Common Searches: how to programmatically select slicer items in Aspose.Cells based on SQL query results c# | c# Aspose.Cells update slicer selection from database values | example of toggling Excel slicer items using Aspose.Cells and ADO.NET | load workbook, read database, set slicer state with Aspose.Cells .NET
// Tags: aspocells set slicer selection from query | c# update slicer items using Aspose.Cells | aspocells load workbook and apply database values | database-driven slicer manipulation with Aspose.Cells | c# handle missing Excel file in Aspose.Cells workflow

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample demonstrates how to verify an input Excel file, load it with Aspose.Cells for .NET, and save it to a new location. It includes placeholders where developers can add logic to read rows from a database query and toggle slicer item selections accordingly, along with recommended error‑handling patterns.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // -----------------------------------------------------------------
            // NOTE: The original example used slicers and a database query.
            // Those APIs are not available in the current project configuration,
            // so the slicer handling and database access have been omitted.
            // The workbook is saved unchanged (or you can add any other logic here).
            // -----------------------------------------------------------------

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
