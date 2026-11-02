// Title: Insert a new row and trigger automatic formula recalculation in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells insert a row and refresh all formulas | How to recalculate dependent formulas after adding rows with Aspose.Cells .NET | Aspose.Cells calculate formulas after inserting rows in an existing workbook | Programmatically add a row and update formulas in Excel using Aspose.Cells C# | Workbook.CalculateFormula usage after row insertion in Aspose.Cells
// Tags: insert row Aspose.Cells C# | calculate formulas workbook Aspose.Cells | refresh dependent formulas after row insertion | load password‑protected Excel Aspose.Cells | save modified workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel file (including password‑protected workbooks), inserts a new row at a specified index, calls Workbook.CalculateFormula to automatically update all dependent formulas, and saves the modified workbook while handling file‑not‑found and Aspose.Cells exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook. If the file is password‑protected, an empty password is supplied.
            var loadOptions = new LoadOptions
            {
                Password = string.Empty
            };
            var workbook = new Workbook(inputPath, loadOptions);

            // Get the first worksheet (adjust index if needed)
            var sheet = workbook.Worksheets[0];

            // Insert a new row at the desired zero‑based index
            int rowIndexToInsert = 5; // example: insert before row 6
            sheet.Cells.InsertRows(rowIndexToInsert, 1);

            // Recalculate all formulas so dependent cells are updated
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
