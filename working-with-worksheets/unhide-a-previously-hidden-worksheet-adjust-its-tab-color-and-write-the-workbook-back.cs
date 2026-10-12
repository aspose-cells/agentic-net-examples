// Title: How to unhide a hidden worksheet, change its tab color, and save the workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an existing .xlsx file using Aspose.Cells, makes a specific hidden sheet visible, sets the sheet tab to LightBlue, and writes the workbook to a new file. | Provide a step‑by‑step example in C# with Aspose.Cells to locate a worksheet by name, change its Visibility property, assign a custom TabColor, and save the updated Excel workbook.
// Common Searches: Aspose.Cells C# reveal hidden worksheet and set tab color | change Excel sheet tab color programmatically using Aspose.Cells .NET | save workbook after modifying sheet visibility with Aspose.Cells | C# example to load workbook, make hidden sheet visible, and apply LightBlue tab color
// Tags: Aspose.Cells reveal hidden sheet C# | Aspose.Cells assign worksheet tab color | Aspose.Cells write updated workbook | C# modify worksheet visibility Aspose | Excel tab color change Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The program loads Sample.xlsx, reveals the worksheet named "HiddenSheet", sets its tab color to LightBlue, and saves the modified workbook as Sample_Modified.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Path to the existing workbook
        string inputPath = @"C:\Temp\Sample.xlsx";

        // Path where the modified workbook will be saved
        string outputPath = @"C:\Temp\Sample_Modified.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Name of the worksheet to unhide (change as needed)
            string sheetName = "HiddenSheet";

            // Get the worksheet by name
            Worksheet sheet = workbook.Worksheets[sheetName];

            if (sheet != null)
            {
                // Unhide the worksheet (set visibility to Visible)
                sheet.IsVisible = true; // Alternative to Visibility property

                // Set the tab color (example: LightBlue)
                sheet.TabColor = Color.LightBlue;
            }
            else
            {
                Console.WriteLine($"Worksheet '{sheetName}' not found.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors (e.g., loading, saving, processing)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
