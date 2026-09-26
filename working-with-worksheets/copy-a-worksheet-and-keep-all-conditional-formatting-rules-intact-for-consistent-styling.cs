// Title: Copy a worksheet in an Excel workbook while preserving all conditional formatting using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to duplicate the worksheet named 'Sheet1' into a new sheet called 'Sheet1_Copy' and ensure every conditional formatting rule is retained. | Programmatically add a blank worksheet to a workbook and copy an existing sheet into it with full formatting, styles, and conditional rules using the Worksheet.Copy method in Aspose.Cells for .NET. | Create a C# script that loads an existing .xlsx file, copies a specific worksheet to a new position in the same workbook, and saves the result without losing any conditional formatting.
// Common Searches: Aspose.Cells C# copy worksheet without losing conditional formatting | How to duplicate an Excel sheet and keep conditional formatting rules using .NET | Worksheet.Copy method preserve conditional formats Aspose.Cells example
// Tags: worksheet.copy preserve conditional formatting aspocells | aspocells duplicate worksheet retain styles | c# copy excel sheet with conditional rules | aspocells copy sheet without losing formatting | excel workbook clone worksheet .net

using System;
using System.IO;
using Aspose.Cells;

// The example checks for source.xlsx, creates it if missing, loads the workbook, adds a new worksheet, copies the existing 'Sheet1' to the new sheet using Worksheet.Copy, renames it to 'Sheet1_Copy', and saves as output.xlsx, preserving all formatting including conditional formatting rules.
class Program
{
    static void Main()
    {
        try
        {
            string sourcePath = "source.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the source file exists; create a simple workbook if it does not.
            if (!File.Exists(sourcePath))
            {
                var tempWb = new Workbook();
                tempWb.Worksheets[0].Name = "Sheet1";
                tempWb.Save(sourcePath);
            }

            // Load the existing workbook.
            Workbook workbook = new Workbook(sourcePath);

            // Retrieve the worksheet to copy.
            Worksheet sourceSheet = workbook.Worksheets["Sheet1"];
            if (sourceSheet == null)
            {
                Console.WriteLine("Worksheet 'Sheet1' not found.");
                return;
            }

            // Add a new empty worksheet and obtain its index.
            int destinationIndex = workbook.Worksheets.Add();

            // Copy the source worksheet (including formatting, styles, etc.) to the new worksheet.
            // Use Worksheet.Copy to avoid ambiguity with System.MemoryExtensions.CopyTo.
            workbook.Worksheets[destinationIndex].Copy(sourceSheet);

            // Optionally rename the copied worksheet.
            workbook.Worksheets[destinationIndex].Name = "Sheet1_Copy";

            // Save the workbook with the copied worksheet.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
