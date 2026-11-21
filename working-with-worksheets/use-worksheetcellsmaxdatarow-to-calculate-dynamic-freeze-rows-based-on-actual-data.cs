// Title: Determine the last populated row with Worksheet.Cells.MaxDataRow and apply dynamic FreezePanes in Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells freeze top rows up to the last data row automatically | how to use MaxDataRow to set FreezePanes in Aspose.Cells .NET | dynamic freeze panes based on actual data range in Excel using Aspose.Cells | retrieve maximum data row index in worksheet with Aspose.Cells C# | apply FreezePanes after calculating max data row in Aspose.Cells workbook
// Tags: Worksheet.Cells.MaxDataRow | dynamic FreezePanes with Aspose.Cells | C# programmatic row freeze in Excel | determine last populated row Aspose.Cells | freeze rows up to max data row .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, uses Worksheet.Cells.MaxDataRow to find the zero‑based index of the last row containing data, freezes all rows above that index with FreezePanes, and saves the updated file.
class DynamicFreezeRows
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine the last row that contains data (zero‑based index)
            int maxDataRow = worksheet.Cells.MaxDataRow;

            // Freeze rows up to the last data row.
            // Use the 4‑parameter overload: FreezePanes(row, column, totalRows, totalColumns)
            // Setting totalRows and totalColumns to 0 freezes everything above the specified row and left of the column.
            worksheet.FreezePanes(maxDataRow + 1, 0, 0, 0);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
