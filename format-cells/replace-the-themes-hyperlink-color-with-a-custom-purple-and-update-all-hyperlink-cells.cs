// Title: Replace the workbook’s theme hyperlink color with a custom purple and apply it to all hyperlink cells using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to set a custom purple RGB font color for every cell that contains a hyperlink in an Excel workbook. | Iterate through each worksheet’s Hyperlink collection and update the Font.Color style of the linked cells to a specific purple shade. | Load an existing .xlsx file, change the theme hyperlink color to a user‑defined purple, and save the modified workbook with Aspose.Cells.
// Common Searches: aspnet change hyperlink font color to purple in Excel using Aspose.Cells | c# programmatically update all hyperlink colors in an existing workbook | how to apply custom theme hyperlink color across all sheets with Aspose.Cells | set RGB color for hyperlink cells in Aspose.Cells .NET example | iterate workbook hyperlinks and modify cell style Aspose.Cells C#
// Tags: set custom hyperlink font color Aspose.Cells C# | update hyperlink cell style across worksheets | apply RGB purple color to Excel hyperlinks | modify theme hyperlink color Aspose.Cells .NET | iterate Hyperlink collection Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, defines a purple RGB color, loops through each worksheet and its Hyperlink collection, updates the font color of every cell within each hyperlink range to the custom color, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Custom hyperlink color (purple)
            Color hyperlinkColor = Color.FromArgb(128, 0, 128);

            // Iterate through each worksheet
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                Cells cells = worksheet.Cells;

                // Process all hyperlinks in the worksheet
                foreach (Hyperlink hyperlink in worksheet.Hyperlinks)
                {
                    CellArea area = hyperlink.Area;

                    for (int row = area.StartRow; row <= area.EndRow; row++)
                    {
                        for (int col = area.StartColumn; col <= area.EndColumn; col++)
                        {
                            Cell cell = cells[row, col];
                            Style style = cell.GetStyle();
                            style.Font.Color = hyperlinkColor; // Apply custom color
                            cell.SetStyle(style);
                        }
                    }
                }
            }

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
