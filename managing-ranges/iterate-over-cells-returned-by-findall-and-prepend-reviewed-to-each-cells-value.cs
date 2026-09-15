// Title: How to prepend "Reviewed:" to every non‑empty cell in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an XLSX workbook with Aspose.Cells, scans the used range of the first worksheet, and prefixes each non‑null cell's string value with "Reviewed:" before saving the file. | Update an existing Aspose.Cells C# program so that it adds the text "Reviewed:" in front of the current content of every populated cell and writes the changes to a new workbook.
// Common Searches: C# Aspose.Cells add prefix to all populated cells in a worksheet | How to update every cell value in an Excel file using Aspose.Cells .NET | Iterate over used range and prepend text to cell values with Aspose.Cells | Bulk prepend string to Excel cells in C# using Aspose.Cells library | Aspose.Cells replace cell text with custom label for non‑empty cells
// Tags: cell value prefix Aspose.Cells C# | iterate used range modify cells Aspose.Cells | bulk update worksheet cells .NET | add custom label to Excel cells Aspose.Cells | process non‑empty cells Aspose.Cells workbook

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX file with Aspose.Cells, accesses the first worksheet, determines its used rows and columns, iterates through each cell in that range, and for every cell that contains data it prefixes the original string with "Reviewed:" before saving the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range of the worksheet
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Iterate through each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // Process only cells that contain a value
                    if (cell != null && cell.Type != CellValueType.IsNull)
                    {
                        string originalValue = cell.StringValue;
                        cell.PutValue("Reviewed:" + originalValue);
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
