// Title: How to unhide a hidden worksheet, change its tab color, and save the workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an existing .xlsx file, makes a hidden worksheet visible, sets its tab color to LightBlue, and saves the workbook. | Provide an Aspose.Cells example that retrieves a worksheet by name, changes its visibility, applies a custom tab color, and writes the updated file.
// Common Searches: how to reveal a hidden sheet and set its tab color using Aspose.Cells for .NET | Aspose.Cells C# change worksheet tab color after unhiding | example code to load workbook, unhide worksheet, set tab color, and save with Aspose.Cells
// Tags: worksheet visibility Aspose.Cells | tab color setting Aspose.Cells | save workbook Aspose.Cells | load workbook from file Aspose.Cells | modify worksheet properties Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads 'input.xlsx', makes the worksheet named 'HiddenSheet' visible, sets its tab color to LightBlue, and saves the result as 'output.xlsx' using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";
                const string sheetName = "HiddenSheet";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the worksheet by name
                Worksheet sheet = workbook.Worksheets[sheetName];
                if (sheet == null)
                {
                    Console.WriteLine($"Worksheet '{sheetName}' not found.");
                    return;
                }

                // Unhide the worksheet
                sheet.IsVisible = true; // alternatively: sheet.Visibility = VisibilityType.Visible;

                // Set the tab color
                sheet.TabColor = Color.LightBlue;

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
