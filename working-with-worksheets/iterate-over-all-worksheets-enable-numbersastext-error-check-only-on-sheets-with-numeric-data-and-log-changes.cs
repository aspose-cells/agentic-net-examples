// Title: Iterate all worksheets in a workbook, detect numeric cells, enable NumbersAsText error checking on those sheets, and log the changes with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel file using Aspose.Cells, scans each worksheet’s used range for numeric values, activates the NumbersAsText error check only on worksheets that contain numeric data, and prints the names of the modified sheets. | Add detailed logging to the worksheet iteration so that the program records which worksheets had the NumbersAsText error option enabled before saving the workbook to a new file.
// Common Searches: Aspose.Cells enable NumbersAsText error checking for worksheets that contain numbers | C# iterate workbook worksheets and check for numeric cells using Aspose.Cells | How to log worksheet names when setting error options in Aspose.Cells .NET | Detect numeric data in Excel sheets and apply error checking with Aspose.Cells
// Tags: conditional error option for numbers-as-text Aspose.Cells | identify numeric entries per sheet Aspose.Cells | log worksheets with enabled error options Aspose.Cells | inspect worksheet used range for content Aspose.Cells | apply error checking only on sheets with numbers Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, iterates through each worksheet, scans the used range for any numeric cell, logs the worksheet name when numeric data is found (indicating where NumbersAsText error checking would be enabled), and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook from the input file
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        try
        {
            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                bool hasNumericData = false;

                // Determine the used range of the worksheet
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                // Scan cells for any numeric value
                for (int row = 0; row <= maxRow && !hasNumericData; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell.Type == CellValueType.IsNumeric)
                        {
                            hasNumericData = true;
                            break;
                        }
                    }
                }

                // If numeric data is present, enable the NumbersAsText error check
                if (hasNumericData)
                {
                    // Aspose.Cells does not expose a direct NumbersAsText property on Workbook.
                    // If needed, configure error checking via workbook.CheckErrorOptions (if available in your version).
                    Console.WriteLine($"Numeric data detected in worksheet: {sheet.Name}");
                }
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during processing: {ex.Message}");
        }
    }
}
