// Title: Update default workbook font and apply a custom heading style in an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: how to change the default font of an Excel workbook with Aspose.Cells in C# | apply a custom bold heading font to the first row of an Excel sheet using Aspose.Cells .NET | set corporate body font for all cells in an existing .xlsx file with Aspose.Cells C# | programmatically modify Excel theme fonts using Aspose.Cells API | Aspose.Cells C# change font for header row only
// Tags: default workbook font Aspose.Cells C# | custom heading style Aspose.Cells | apply corporate typeface to Excel workbook | set bold header font programmatically | update Excel theme fonts using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The code loads an existing .xlsx workbook, changes the workbook's default body font to a corporate typeface, creates a bold heading style with a separate corporate font, applies that style to the first row, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Update the default font for the workbook (as Theme API may not be available in all versions)
                // This sets the font used for cells that do not have an explicit style applied.
                workbook.DefaultStyle.Font.Name = "CorporateBodyFont";

                // Optionally, you can update the style of the first worksheet's cells to use a heading font.
                // Here we demonstrate setting the first row (assumed headings) to a different font.
                Worksheet sheet = workbook.Worksheets[0];
                Style headingStyle = workbook.CreateStyle();
                headingStyle.Font.Name = "CorporateHeadingFont";
                headingStyle.Font.IsBold = true;

                // Apply the heading style to the first row (A1 to the last used column in row 1)
                int lastColumn = sheet.Cells.MaxColumn;
                for (int col = 0; col <= lastColumn; col++)
                {
                    Cell cell = sheet.Cells[0, col];
                    cell.SetStyle(headingStyle);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
