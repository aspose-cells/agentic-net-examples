// Title: Load an existing Excel workbook and set a SUM formula in cell B2 using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load existing workbook and add SUM formula to B2 | set =SUM(A1:A10) formula in cell B2 with Aspose.Cells .NET | apply Cell.Formula to a cell after opening an Excel file using Aspose.Cells C# | create workbook with sample data when input.xlsx is missing Aspose.Cells example | handle password‑protected workbook while assigning formulas with Aspose.Cells
// Tags: load workbook with Aspose.Cells C# | assign SUM formula using Cell.Formula | programmatically write Excel formula Aspose.Cells | create workbook when file not found Aspose.Cells | handle password protected workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example opens "input.xlsx" with Aspose.Cells; if the file does not exist it creates a new workbook and fills cells A1‑A10 with values 1‑10. It then retrieves the first worksheet, assigns the formula "=SUM(A1:A10)" to cell B2 via the Cell.Formula property, and saves the result as "output.xlsx". The code includes handling for password‑protected workbooks and generic exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            if (File.Exists(inputPath))
            {
                // Load existing workbook (no password)
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook if the input file is missing
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                // Populate A1:A10 with sample values
                for (int i = 0; i < 10; i++)
                {
                    ws.Cells[i, 0].PutValue(i + 1);
                }
            }
        }
        catch (CellsException ex)
        {
            // Handle password‑protected workbook or other Cells exceptions
            if (ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("The workbook is password protected and cannot be opened without a password.");
            }
            else
            {
                Console.WriteLine($"Cells error loading workbook: {ex.Message}");
            }
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Assign a SUM formula to cell B2
            Cell targetCell = worksheet.Cells["B2"];
            targetCell.Formula = "=SUM(A1:A10)";

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }
}
