// Title: Convert HTML to XLSX and apply custom scientific notation (0.00E+00) to numeric cells using Aspose.Cells for .NET
// AI Prompts: Load an HTML file into an Aspose.Cells Workbook and save the workbook as an XLSX file. | Iterate over the worksheet's used range, convert any string that represents a number to a double, and set the custom format "0.00E+00" for cells whose absolute value is below 0.001 or above 1,000,000. | Create the output directory if it does not exist before writing the formatted workbook.
// Common Searches: how to import HTML into Excel with Aspose.Cells and keep scientific notation | C# Aspose.Cells format numbers as scientific notation after loading HTML | convert numeric strings to numbers when converting HTML to XLSX using Aspose.Cells | apply scientific notation to cells when converting HTML to Excel | Aspose.Cells load HTML and automatically format very large or very small values
// Tags: HTML to XLSX conversion Aspose.Cells | custom scientific notation format Aspose.Cells | numeric string to double conversion C# | apply custom number format to used range | save workbook as Xlsx with formatting

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using CellsRange = Aspose.Cells.Range;

// The example loads an HTML file into an Aspose.Cells Workbook, scans the used cell range, converts string values that represent numbers into numeric cells, applies the custom scientific notation format "0.00E+00" to values that are very small or very large, and saves the result as an XLSX workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.html";
            const string outputPath = "output.xlsx";

            // Verify that the input HTML file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the HTML file into a new Workbook using HtmlLoadOptions
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Get the first worksheet (HTML is usually loaded into the first sheet)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define a custom number format for scientific notation
            const string scientificFormat = "0.00E+00";

            // Determine the used range of the worksheet
            CellsRange usedRange = cells.MaxDisplayRange;
            int startRow = usedRange.FirstRow;
            int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
            int startCol = usedRange.FirstColumn;
            int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

            // Iterate through all cells in the used range
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    Cell cell = cells[row, col];

                    // If the cell already contains a numeric value
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        double value = cell.DoubleValue;

                        // Apply scientific format for very small or very large numbers
                        if (value != 0 && (Math.Abs(value) < 0.001 || Math.Abs(value) > 1_000_000))
                        {
                            Style style = cell.GetStyle();
                            style.Custom = scientificFormat;
                            cell.SetStyle(style);
                        }
                    }
                    // If the cell contains a string that can be parsed as a number
                    else if (cell.Type == CellValueType.IsString)
                    {
                        if (double.TryParse(cell.StringValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed))
                        {
                            // Replace the string with the numeric value
                            cell.PutValue(parsed);

                            // Apply the scientific format
                            Style style = cell.GetStyle();
                            style.Custom = scientificFormat;
                            cell.SetStyle(style);
                        }
                    }
                }
            }

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an Excel file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
