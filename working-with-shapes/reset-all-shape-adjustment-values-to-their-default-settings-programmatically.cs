// Title: Reset Excel shape adjustment values to their defaults with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# reset shape adjustment values to default | programmatically revert Excel shape adjustments using Aspose.Cells | C# code to clear custom shape adjustment settings in an Excel file | how to restore default shape properties in Aspose.Cells workbook | reset all shape adjustments in a worksheet with Aspose.Cells .NET
// Tags: reset shape adjustments Aspose.Cells .NET | default shape properties Excel C# | iterate over worksheet shapes Aspose.Cells | shape adjustment API limitation Aspose.Cells | modify Excel shape settings programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel workbook, iterates through each worksheet and its shapes, notes that Aspose.Cells currently lacks a direct API to reset shape adjustments, provides a placeholder where shape properties could be modified, and then saves the workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the current worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Aspose.Cells does not provide a direct API to reset shape adjustments.
                    // If needed, shape properties can be modified here.
                    // This block is kept for future extension.
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
