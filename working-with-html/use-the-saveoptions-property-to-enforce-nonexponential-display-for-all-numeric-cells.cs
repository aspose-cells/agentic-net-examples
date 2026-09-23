// Title: How to disable scientific notation for all numeric cells when saving an XLSX workbook with Aspose.Cells for .NET
// AI Prompts: Write C# code that loops through every worksheet and cell in an Aspose.Cells workbook, applies a custom number format to numeric cells to prevent scientific notation, and saves the file as XLSX. | Show how to configure Aspose.Cells SaveOptions so that all numeric values are written without exponential notation in the output workbook. | Provide a complete example that creates a placeholder Excel file, formats all numeric cells with a plain number pattern, and saves the workbook using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells .NET prevent scientific notation when exporting to XLSX | C# apply plain number format to all cells in an Excel workbook using Aspose.Cells | How to force non‑exponential display for numeric values in Aspose.Cells SaveOptions | Iterate over cells and set custom number format in Aspose.Cells C# example | Disable exponential notation for numbers in saved Excel file with Aspose.Cells
// Tags: apply custom numeric style Aspose.Cells | disable exponential notation numeric cells .NET | iterate worksheets and set cell style Aspose | save workbook with non‑scientific numeric format XLSX | Aspose.Cells numeric formatting via SaveOptions

using Aspose.Cells;
using System;
using System.IO;

// The sample loads (or creates) an Excel workbook, iterates through each worksheet and cell, applies a custom number format "0.####################" to numeric cells to force plain (non‑exponential) display, and saves the result as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Ensure the input file exists; create a simple workbook if missing
        if (!File.Exists(inputPath))
        {
            try
            {
                Workbook tempWb = new Workbook();
                tempWb.Worksheets[0].Name = "Sheet1";
                tempWb.Save(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to create placeholder input file: {ex.Message}");
                return;
            }
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Apply only to numeric cells
                        if (cell.Type == CellValueType.IsNumeric)
                        {
                            // Use a custom number format that forces plain (non‑exponential) display
                            Style style = cell.GetStyle();
                            style.Custom = "0.####################";
                            cell.SetStyle(style);
                        }
                    }
                }
            }

            // Save the workbook in XLSX format
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
