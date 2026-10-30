// Title: Insert a row and automatically recalculate dependent formulas in an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells insert row and refresh formulas | How to trigger formula recalculation after adding rows in Aspose.Cells | Aspose.Cells Workbook.CalculateFormula example after row insertion | Update dependent formulas automatically when inserting rows with Aspose.Cells | Recalculate all workbook formulas after inserting a new row in .NET
// Tags: insert rows Aspose.Cells C# | calculate all formulas Aspose.Cells | recalculate formulas for entire workbook Aspose.Cells | auto update dependent formulas Aspose.Cells | Excel row insertion formula refresh .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file or creates a new workbook, ensures a worksheet exists, inserts a row at a specified index, calls Workbook.CalculateFormula to refresh all formulas affected by the insertion, and saves the updated file, handling missing files and password‑protected workbooks gracefully.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            // Load existing workbook if it exists and is not password‑protected
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (CellsException ex) when (ex.Message.Contains("Please provide password"))
                {
                    // Password required – create a new workbook instead
                    workbook = new Workbook();
                }
            }
            else
            {
                // Input file not found – start with a new workbook
                workbook = new Workbook();
            }

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
                workbook.Worksheets.Add();

            Worksheet sheet = workbook.Worksheets[0];

            // Insert a new row at the desired index (zero‑based)
            int insertRowIndex = 5;   // adjust as needed
            int numberOfRows = 1;     // number of rows to insert
            sheet.Cells.InsertRows(insertRowIndex, numberOfRows);

            // Recalculate all formulas after row insertion
            workbook.CalculateFormula();

            // Save the updated workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
