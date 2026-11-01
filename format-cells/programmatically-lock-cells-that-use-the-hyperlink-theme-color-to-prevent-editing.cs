// Title: Lock cells that use the default hyperlink blue color in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells lock cells with hyperlink blue font color in C# | C# protect worksheet and lock cells based on font color using Aspose.Cells | How to detect and lock hyperlink‑styled cells in an Excel file with Aspose.Cells | Set IsLocked for cells with RGB 0,0,255 font using Aspose.Cells .NET | Programmatically protect Excel cells that use the hyperlink theme color in C#
// Tags: lock cells by font color Aspose.Cells | worksheet protection with cell style evaluation | detect hyperlink theme color C# Aspose.Cells | set cell IsLocked programmatically | Excel workbook cell style locking Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example loads a workbook, protects each worksheet, scans all used cells, identifies those whose font color matches the standard hyperlink blue (RGB 0,0,255), sets their style IsLocked flag to true, and saves the updated file.
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
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Process each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Protect the worksheet (default protection type)
                    sheet.Protect(ProtectionType.All);

                    // Determine the used range of the sheet
                    int maxRow = sheet.Cells.MaxDataRow;
                    int maxCol = sheet.Cells.MaxDataColumn;

                    // Define the typical hyperlink blue color (RGB 0,0,255)
                    Color hyperlinkColor = Color.FromArgb(0, 0, 255);
                    int hyperlinkArgb = hyperlinkColor.ToArgb();

                    // Iterate through all used cells
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];
                            Style style = cell.GetStyle();

                            // Check if the font color matches the typical hyperlink color
                            if (style.Font.Color.ToArgb() == hyperlinkArgb)
                            {
                                // Lock the cell to prevent editing
                                style.IsLocked = true;
                                cell.SetStyle(style);
                            }
                        }
                    }
                }
                catch (Exception exSheet)
                {
                    Console.WriteLine($"Error processing sheet '{sheet.Name}': {exSheet.Message}");
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception exSave)
            {
                Console.WriteLine($"Error saving workbook: {exSave.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
