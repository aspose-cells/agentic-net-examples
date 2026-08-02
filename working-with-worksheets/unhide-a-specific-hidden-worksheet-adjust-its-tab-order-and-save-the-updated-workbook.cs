// Title: Aspose.Cells for .NET – Unhide a Worksheet, Move It to the First Tab, and Save the Workbook (C#)
// Description: C# example that loads an existing XLSX file with Aspose.Cells, makes a hidden worksheet visible, moves it to the first tab (index 0), and saves the updated workbook to a new file.
// Keywords: Aspose.Cells unhide worksheet C# | move worksheet to first tab Aspose.Cells | set worksheet visibility Aspose.Cells .NET | save workbook after reordering sheets | C# Aspose.Cells worksheet order | Aspose.Cells hidden sheet handling
// Common Searches: how to unhide a specific worksheet using Aspose.Cells for .NET | C# move hidden sheet to first tab Aspose.Cells | Aspose.Cells change worksheet visibility and reorder tabs | unhide and reposition worksheet Aspose.Cells example
// Developer Intent: Make a hidden worksheet visible, place it as the first tab in the workbook, and persist the changes to a new file using Aspose.Cells for .NET.
// Use Cases: Expose a confidential sheet before publishing the workbook to end users. | Ensure a summary or cover sheet appears as the first tab in generated reports. | Automate workbook cleanup by unhiding hidden tabs and positioning them at the beginning prior to distribution.
// AI Prompts: Generate C# code with Aspose.Cells to unhide a worksheet named "Report" and move it to the second tab position. | Show robust error‑handling for loading a workbook, verifying a worksheet exists, changing its visibility, reordering tabs, and saving the file with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace MyApp
{
    // C# example that loads an existing XLSX file with Aspose.Cells, makes a hidden worksheet visible, moves it to the first tab (index 0), and saves the updated workbook to a new file.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";
                string hiddenSheetName = "HiddenSheet";

                // Ensure the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the worksheet by name
                Worksheet sheet = workbook.Worksheets[hiddenSheetName];
                if (sheet == null)
                {
                    Console.WriteLine($"Worksheet '{hiddenSheetName}' not found.");
                    return;
                }

                // Unhide the worksheet
                sheet.IsVisible = true;

                // Move the worksheet to the first tab position
                sheet.MoveTo(0);

                // Save the updated workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
