// Title: Extract and save formula text from a specific cell range in an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells get formula strings from a rectangular range of cells | how to list formulas of cells A1:C10 with Aspose.Cells .NET | extract Excel cell formulas and write them to a new worksheet using Aspose.Cells | retrieve formula or displayed value for each cell in a range with Aspose.Cells | bulk export formulas from an Excel sheet to another sheet in C#
// Tags: Aspose.Cells extract formula text C# | read cell Formula property Aspose.Cells | write extracted formulas to new worksheet Aspose.Cells | bulk retrieve Excel formulas .NET | convert column index to name Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, defines a rectangular range (first 10 rows and columns A‑C), iterates through each cell to capture its Formula property or displayed value, stores the results in a 2‑D array, prints each address with the extracted text, writes the collected strings to a newly added worksheet named "ExtractedFormulas", and saves the workbook as a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range to extract formulas (zero‑based indices)
            int firstRow = 0;          // Row 1
            int firstColumn = 0;       // Column A
            int totalRows = 10;        // Number of rows to include
            int totalColumns = 3;      // Number of columns to include (A, B, C)

            // Prepare a 2‑D array for formula texts
            string[,] formulaTexts = new string[totalRows, totalColumns];

            // Retrieve formula text for each cell in the range
            for (int i = 0; i < totalRows; i++)
            {
                for (int j = 0; j < totalColumns; j++)
                {
                    Cell cell = worksheet.Cells[firstRow + i, firstColumn + j];
                    // If the cell contains a formula, get its formula; otherwise, get the displayed value
                    formulaTexts[i, j] = !string.IsNullOrEmpty(cell.Formula) ? cell.Formula : cell.StringValue;
                }
            }

            // Output each cell address with its formula text
            for (int i = 0; i < totalRows; i++)
            {
                for (int j = 0; j < totalColumns; j++)
                {
                    // Convert column index to letter (A, B, …)
                    string columnLetter = CellsHelper.ColumnIndexToName(firstColumn + j);
                    int rowNumber = firstRow + i + 1; // 1‑based row number for display
                    Console.WriteLine($"Cell {columnLetter}{rowNumber}: {formulaTexts[i, j]}");
                }
            }

            // Optional: write the extracted formulas to a new worksheet
            Worksheet resultSheet = workbook.Worksheets.Add("ExtractedFormulas");
            for (int i = 0; i < totalRows; i++)
            {
                for (int j = 0; j < totalColumns; j++)
                {
                    resultSheet.Cells[i, j].PutValue(formulaTexts[i, j]);
                }
            }

            // Save the workbook with the new sheet
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
