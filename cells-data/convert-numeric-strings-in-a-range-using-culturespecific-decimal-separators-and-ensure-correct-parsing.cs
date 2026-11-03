// Title: Convert string cells to numbers in a specific Excel range using a French decimal separator with Aspose.Cells for .NET
// Common Searches: c# aspnet convert excel string cells to double using fr-FR culture asp osee.cells | parse numeric strings with comma decimal separator in Aspose.Cells range | replace string values with numbers in a specific worksheet range using Aspose.Cells .NET | Aspose.Cells read Excel cells as strings then convert to numeric based on culture | set culture for numeric parsing when processing Excel files with Aspose.Cells
// Tags: culture-aware numeric parsing Aspose.Cells | Excel range iteration for value conversion .NET | French locale decimal handling in Excel | Aspose.Cells cell type conversion from string to double | Workbook save after cell value transformation

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads an XLSX workbook, creates a defined range (e.g., A1:B10), iterates each cell, and for cells containing string values attempts to parse them as doubles using the French (fr-FR) culture, which uses a comma as the decimal separator. Successfully parsed values replace the original strings, and the workbook is saved to the output path.
class NumericStringConverter
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = @"C:\Data\Input.xlsx";
        string outputPath = @"C:\Data\Output.xlsx";

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        // Ensure output directory exists
        string outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        // Culture that uses the desired decimal separator (e.g., French uses comma)
        CultureInfo targetCulture = new CultureInfo("fr-FR");

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define the range that contains numeric strings (adjust address as required)
            string rangeAddress = "A1:B10";

            // Create the range (use alias to avoid conflict with System.Range)
            AsposeRange range = cells.CreateRange(rangeAddress);

            // Determine the bounds of the range
            int startRow = range.FirstRow;
            int endRow = range.FirstRow + range.RowCount - 1;
            int startCol = range.FirstColumn;
            int endCol = range.FirstColumn + range.ColumnCount - 1;

            // Iterate through each cell in the range
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only cells that currently hold string values
                    if (cell.Type == CellValueType.IsString)
                    {
                        string text = cell.StringValue.Trim();

                        // Try to parse the string using the target culture's decimal separator
                        if (double.TryParse(text, NumberStyles.Any, targetCulture, out double numericValue))
                        {
                            // Replace the string with the parsed numeric value
                            cell.PutValue(numericValue);
                        }
                        // If parsing fails, the cell remains unchanged
                    }
                }
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
