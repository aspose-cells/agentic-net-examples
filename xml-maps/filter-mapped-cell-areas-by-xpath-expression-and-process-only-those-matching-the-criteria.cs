// Title: How to prepend a custom prefix to all string cells in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells prepend text to string cells in a worksheet | Iterate over used range and modify only string values with Aspose.Cells .NET | Add a prefix to Excel cell contents using Aspose.Cells in C# | Update only string-type cells in an Excel file with Aspose.Cells API | Save workbook after changing cell text with Aspose.Cells for .NET | Check cell value type before writing new data using Aspose.Cells C#
// Tags: prepend prefix to string cells Aspose.Cells C# | iterate used cells range Aspose.Cells .NET | modify cell value based on type Aspose.Cells | save updated workbook Aspose.Cells C# | process string cell values Excel Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, iterates through every used cell in the first worksheet, detects cells containing string values, prepends a "Processed: " prefix to each, and then saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Iterate through all used cells.
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Example processing: if the cell contains a string, prepend a marker.
                    if (cell.Type == CellValueType.IsString)
                    {
                        string original = cell.StringValue;
                        cell.PutValue("Processed: " + original);
                    }

                    // Additional processing logic can be added here.
                }
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
