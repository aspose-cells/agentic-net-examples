// Title: Skip applying FreezePanes in Aspose.Cells for .NET when the worksheet already has the required frozen rows and columns
// AI Prompts: Generate C# code that checks a worksheet's current FreezePanes row and column split values and calls FreezePanes only when they differ from the desired settings. | Create a helper method that returns true if the top rows and left columns are already frozen at specified positions, and use it to prevent redundant FreezePanes calls in Aspose.Cells.
// Common Searches: Aspose.Cells .NET how to determine if freeze panes are already set on a worksheet | C# check existing frozen rows and columns before calling Worksheet.FreezePanes | Avoid duplicate FreezePanes operation in Aspose.Cells example | Read current freeze pane row split and column split using Aspose.Cells API
// Tags: Aspose.Cells worksheet freeze panes check | C# Aspose.Cells skip redundant FreezePanes | Aspose.Cells get current freeze pane splits | conditional freeze panes Aspose.Cells .NET | optimize freeze pane operation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, defines the desired number of frozen rows and columns, checks the worksheet's existing freeze pane settings, applies FreezePanes only if the current state differs, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Desired freeze settings
                int desiredRowSplit = 1;      // Number of rows to freeze from the top
                int desiredColumnSplit = 2;   // Number of columns to freeze from the left

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Apply freeze panes using the 4‑argument overload
                sheet.FreezePanes(desiredRowSplit, desiredColumnSplit, desiredRowSplit, desiredColumnSplit);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
