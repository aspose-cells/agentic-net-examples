// Title: Programmatically lock Excel cells styled with the Hyperlink theme color using Aspose.Cells for .NET
// AI Prompts: Identify cells whose font color matches the default hyperlink blue, set their Style.IsLocked property to true, and protect the worksheet with a password using Aspose.Cells. | Iterate over the used range of each worksheet, apply a locked style to cells using the Hyperlink theme color, then save the workbook as a protected file.
// Common Searches: Aspose.Cells C# lock cells with hyperlink font color and protect worksheet | How to protect Excel cells that use the default hyperlink blue using Aspose.Cells .NET | Set IsLocked for cells styled with Hyperlink theme in Aspose.Cells workbook | C# code to lock cells based on font color and save protected XLSX with Aspose.Cells
// Tags: hyperlink theme color cell locking Aspose.Cells | worksheet protection after applying style lock .NET | detect and lock default hyperlink blue cells C# | used range style modification Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The example loads an existing XLSX workbook, scans each worksheet's used range, locks any cell whose font color equals the default hyperlink blue, protects the sheet with a password, and saves the result as a new protected file.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = @"C:\Input\Sample.xlsx";
            string outputPath = @"C:\Output\Sample_Locked.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the worksheet
                var usedRange = sheet.Cells.MaxDisplayRange; // Aspose.Cells.Range

                // Determine start/end rows and columns
                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startColumn = usedRange.FirstColumn;
                int endColumn = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                // Loop through each cell in the used range
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startColumn; col <= endColumn; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        Style style = cell.GetStyle();

                        // Lock cells whose font color is the default hyperlink blue
                        if (style.Font.Color.ToArgb() == Color.Blue.ToArgb())
                        {
                            style.IsLocked = true;
                            cell.SetStyle(style);
                        }
                    }
                }

                // Protect the worksheet with a password
                sheet.Protect(ProtectionType.All, "password", string.Empty);
            }

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
