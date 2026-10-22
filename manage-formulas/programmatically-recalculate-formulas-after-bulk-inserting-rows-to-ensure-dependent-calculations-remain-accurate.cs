// Title: Recalculate workbook formulas after inserting multiple rows with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# recalculate all formulas after inserting rows | How to refresh Excel calculations after bulk row insertion using Aspose.Cells | Calculate dependent formulas in a workbook after adding multiple rows in .NET | InsertRows method formula update Aspose.Cells example
// Tags: Worksheet.InsertRows Aspose.Cells | Workbook.CalculateFormula .NET | bulk row insertion Excel Aspose.Cells | refresh dependent formulas after row insert | C# update Excel calculations Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file (or creates a new workbook if missing or password‑protected), inserts ten rows starting at row index 5 on the first worksheet, calls Workbook.CalculateFormula() to refresh all dependent formulas, and saves the updated workbook to output.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        // Load existing workbook if it exists; otherwise create a new one
        if (File.Exists(inputPath))
        {
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (CellsException ex) when (ex.Message.Contains("password"))
            {
                // Workbook is password‑protected; cannot open without password.
                // Fallback to a new workbook.
                Console.WriteLine("The input workbook is password protected. A new workbook will be created.");
                workbook = new Workbook();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }
        }
        else
        {
            workbook = new Workbook(); // creates a default workbook with one worksheet
        }

        try
        {
            // Get the first worksheet (adjust index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Insert 10 rows starting at row index 5 (zero‑based, i.e., after row 5)
            int startRow = 5;
            int rowsToInsert = 10;
            worksheet.Cells.InsertRows(startRow, rowsToInsert);

            // Recalculate all formulas to reflect the inserted rows
            workbook.CalculateFormula();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Runtime error: {ex.Message}");
        }
    }
}
