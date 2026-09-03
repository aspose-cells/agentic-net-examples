// Title: Copy a worksheet from an existing Excel file to a new workbook and freeze the header row with Aspose.Cells for .NET (C#)
// AI Prompts: Load source.xlsx, add a new worksheet to a fresh workbook, copy the first sheet while preserving its name, and apply FreezePanes to the first row using Aspose.Cells in C#. | Create a destination workbook, import a worksheet from another workbook, optionally rename it, and set FreezePanes on the top row programmatically with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# copy worksheet from one workbook to another and keep original sheet name | How to freeze the first row after copying a sheet with Aspose.Cells for .NET | C# example for copying a worksheet and applying FreezePanes using Aspose.Cells | Copy sheet and set FreezePanes on header row in Aspose.Cells .NET library
// Tags: copy worksheet between workbooks Aspose.Cells C# | freeze panes on copied sheet Aspose.Cells | preserve original sheet name Aspose.Cells | load and save Excel files Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads source.xlsx, creates a new workbook, copies the first worksheet into it while preserving the original sheet name, freezes the first row of the copied sheet, and saves the result as output.xlsx, with error handling for missing files.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new destination workbook (contains a default worksheet)
                Workbook destWorkbook = new Workbook();

                // Verify that the source file exists before loading
                string sourcePath = "source.xlsx";
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook
                Workbook srcWorkbook = new Workbook(sourcePath);

                // Get the first worksheet from the source workbook
                Worksheet srcSheet = srcWorkbook.Worksheets[0];

                // Add a new worksheet to the destination workbook and copy the source sheet into it
                int newSheetIndex = destWorkbook.Worksheets.Add();
                Worksheet destSheet = destWorkbook.Worksheets[newSheetIndex];
                destSheet.Copy(srcSheet);
                destSheet.Name = srcSheet.Name; // Preserve original name

                // Freeze the header row (first row) in the copied worksheet
                destSheet.FreezePanes(0, 0, 1, 0);

                // Save the destination workbook
                string outputPath = "output.xlsx";
                destWorkbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
