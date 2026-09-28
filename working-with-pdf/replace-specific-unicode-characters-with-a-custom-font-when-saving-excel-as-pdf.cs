// Title: Replace left and right single‑quote Unicode characters with Arial Unicode MS before converting an Excel workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Locate cells that contain the Unicode characters ‘\u2018’ or ‘\u2019’, change their font to Arial Unicode MS, and then save the workbook as a PDF with Aspose.Cells. | Loop through every worksheet, apply Arial Unicode MS to any string cell that includes left or right single quotation marks, and export the result to PDF.
// Common Searches: Aspose.Cells C# change font for cells containing Unicode left single quote before PDF export | how to apply Arial Unicode MS to specific characters in Excel when saving as PDF using Aspose.Cells | replace Unicode quotation marks with a different font during Excel to PDF conversion in .NET | detect cells with \u2018 or \u2019 and set a custom font in Aspose.Cells before generating PDF | C# Aspose.Cells map special Unicode characters to Arial Unicode MS for PDF output
// Tags: replace Unicode characters with Arial Unicode MS Aspose.Cells | conditional cell font assignment C# | Excel to PDF custom font mapping Aspose.Cells | targeted cell font change .NET | detect left right single quotes worksheet Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook, scans each worksheet for cells containing left or right single‑quote Unicode characters, applies Arial Unicode MS to those cells, and saves the workbook as a PDF.
class ReplaceUnicodeAndSavePdf
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Unicode characters that need a custom font (e.g., left/right single quotes)
            HashSet<char> targetChars = new HashSet<char> { '\u2018', '\u2019' };

            // Custom font name to apply
            const string customFontName = "Arial Unicode MS";

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the worksheet (fully qualified to avoid ambiguity)
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                // Iterate through each cell in the used range
                foreach (Cell cell in usedRange)
                {
                    // Process only string cells
                    if (cell.Type != CellValueType.IsString)
                        continue;

                    string cellText = cell.StringValue;
                    if (string.IsNullOrEmpty(cellText))
                        continue;

                    // Check if the cell contains any target Unicode characters
                    bool containsTarget = false;
                    foreach (char ch in cellText)
                    {
                        if (targetChars.Contains(ch))
                        {
                            containsTarget = true;
                            break;
                        }
                    }

                    if (containsTarget)
                    {
                        try
                        {
                            // Apply the custom font to the whole cell (simplified approach)
                            Style style = cell.GetStyle();
                            style.Font.Name = customFontName;
                            cell.SetStyle(style);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to set font for cell {cell.Name}: {ex.Message}");
                        }
                    }
                }
            }

            // Save the workbook as PDF
            try
            {
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"PDF saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save PDF: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
