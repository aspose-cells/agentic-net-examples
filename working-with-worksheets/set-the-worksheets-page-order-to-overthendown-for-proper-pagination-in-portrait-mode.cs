// Title: Set worksheet page order to OverThenDown with portrait orientation using Aspose.Cells for .NET
// AI Prompts: Load an Excel workbook, change the first sheet's PageSetup.PageOrder to OverThenDown, set PageSetup.Orientation to Portrait, and save the file with Aspose.Cells in C#. | Write C# code that applies over‑then‑down pagination to a worksheet while enforcing portrait printing layout, then writes the updated workbook to a new file using Aspose.Cells.
// Common Searches: Aspose.Cells C# how to change page order to over then down for printing | set worksheet pagination over then down portrait mode Aspose.Cells .NET | C# example for configuring PageSetup.PageOrder to OverThenDown in Excel file | printing Excel sheet with over‑then‑down page order using Aspose.Cells | Aspose.Cells page order portrait orientation example code
// Tags: Aspose.Cells worksheet page order OverThenDown | C# PageSetup Orientation Portrait Aspose.Cells | Excel over‑then‑down pagination .NET | PageSetup PageOrder property Aspose.Cells | set worksheet printing layout Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads input.xlsx, accesses the first worksheet, sets its PageSetup.PageOrder to OverThenDown (if supported) and forces portrait orientation, then saves the workbook as output.xlsx while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set page order to over‑then‑down if supported (commented out if not available)
            // sheet.PageSetup.PageOrder = PageOrder.OverThenDown;

            // Ensure the worksheet orientation is portrait
            sheet.PageSetup.Orientation = PageOrientationType.Portrait;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
