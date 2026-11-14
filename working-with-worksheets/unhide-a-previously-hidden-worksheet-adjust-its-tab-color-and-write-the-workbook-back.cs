// Title: Unhide a hidden worksheet, set its tab color, and save the workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# code to unhide a hidden Excel sheet and change its tab color using Aspose.Cells | Aspose.Cells example for making a worksheet visible and setting tab color before saving | How to locate and unhide the first hidden worksheet when a named sheet is missing in Aspose.Cells | Saving an updated workbook to a new file after modifying worksheet visibility with Aspose.Cells C# | Handling missing worksheet errors while un-hiding and recoloring tabs in Aspose.Cells
// Tags: unhide hidden worksheet Aspose.Cells C# | change worksheet tab color Aspose.Cells | save modified Excel workbook Aspose.Cells | load workbook from file path Aspose.Cells | fallback to first hidden sheet Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

namespace AsposeCellsExample
{
    // The sample loads an existing Excel file, tries to find a worksheet named "HiddenSheet" or the first hidden sheet, makes it visible, sets its tab color to LightBlue, and saves the workbook to a new file, with file‑existence checks and exception handling.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string targetSheetName = "HiddenSheet"; // change as needed

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

                // Try to get the worksheet by name
                Worksheet sheet = workbook.Worksheets[targetSheetName];

                // If not found, attempt to locate the first hidden worksheet
                if (sheet == null)
                {
                    foreach (Worksheet ws in workbook.Worksheets)
                    {
                        if (!ws.IsVisible)
                        {
                            sheet = ws;
                            Console.WriteLine($"Worksheet \"{targetSheetName}\" not found. Using hidden sheet \"{ws.Name}\" instead.");
                            break;
                        }
                    }
                }

                // If still null, there is no worksheet to unhide
                if (sheet == null)
                {
                    Console.WriteLine("No worksheet found to unhide.");
                    return;
                }

                // Unhide the worksheet
                sheet.IsVisible = true;

                // Set the tab color (example: LightBlue)
                sheet.TabColor = Color.LightBlue;

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
