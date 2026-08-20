// Title: Unhide Hidden Worksheet, Change Tab Color, and Save Workbook with Aspose.Cells for .NET (C#)
// Description: C# example that loads an Excel file using Aspose.Cells, locates a worksheet named "HiddenSheet" or the first hidden sheet, makes it visible, sets its tab color to green, and saves the updated workbook to a new file.
// Keywords: Aspose.Cells unhide worksheet C# | set worksheet tab color Aspose.Cells | Excel hidden sheet visibility .NET | save workbook after modifying sheet Aspose | C# find first hidden worksheet Aspose.Cells | Aspose.Cells workbook save example
// Common Searches: how to unhide a hidden worksheet with Aspose.Cells | change Excel sheet tab color using C# Aspose.Cells | save modified workbook after changing sheet properties | C# code to locate first hidden sheet in Excel file | Aspose.Cells example for worksheet visibility and tab color
// Developer Intent: Reveal a hidden worksheet, apply a custom tab color, and write the modified workbook back to disk.
// Use Cases: Programmatically unhide a specific sheet by name before publishing a report. | Automatically detect and reveal the first hidden sheet in legacy workbooks. | Apply corporate color schemes to worksheet tabs after making them visible. | Batch‑process Excel files to ensure all sheets are visible and correctly colored before distribution.
// AI Prompts: Generate C# code with Aspose.Cells that unhides a worksheet called "Report" and sets its tab color to blue, then saves the workbook. | Explain how to iterate through a workbook's worksheets in Aspose.Cells to find the first hidden sheet, make it visible, assign a custom tab color, and persist the changes. | Provide a step‑by‑step guide for handling missing input files when unhide‑and‑color operations fail in Aspose.Cells.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWorksheetUnhideExample
{
    // C# example that loads an Excel file using Aspose.Cells, locates a worksheet named "HiddenSheet" or the first hidden sheet, makes it visible, sets its tab color to green, and saves the updated workbook to a new file.
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
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Try to get the worksheet named "HiddenSheet"
                Worksheet sheet = workbook.Worksheets["HiddenSheet"];

                // If not found, look for the first hidden worksheet
                if (sheet == null)
                {
                    foreach (Worksheet ws in workbook.Worksheets)
                    {
                        if (!ws.IsVisible)
                        {
                            sheet = ws;
                            break;
                        }
                    }
                }

                // If still not found, report and exit
                if (sheet == null)
                {
                    Console.WriteLine("No hidden worksheet found in the workbook.");
                    return;
                }

                // Unhide the worksheet
                sheet.IsVisible = true;

                // Change the tab color
                sheet.TabColor = Color.Green;

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
