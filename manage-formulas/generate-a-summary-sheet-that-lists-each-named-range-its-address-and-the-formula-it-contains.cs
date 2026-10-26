// Title: Generate a summary sheet of all named ranges with their addresses and first‑cell formulas using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# list named ranges and export to new worksheet | how to write named range address and formula to summary sheet with Aspose.Cells | C# code to iterate workbook names collection and capture formulas using Aspose.Cells | create a summary tab that shows each named range's reference and formula in Excel via Aspose.Cells | retrieve first cell formula from named range in Aspose.Cells .NET example
// Tags: enumerate named ranges Aspose.Cells | write named range details to worksheet C# | extract formula from first cell Aspose.Cells | add summary worksheet Aspose.Cells .NET | named range address retrieval Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads an existing workbook (or creates a new one), adds a "Summary" worksheet, writes column headers, then loops through the workbook's named ranges collection. For each name it records the range's name, its reference address, and, if the first cell contains a formula, the formula text. All gathered information is written to the summary sheet and the workbook is saved.
class Program
{
    static void Main()
    {
        // Define input and output file paths
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            // Load the existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook with a default worksheet
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Add a new worksheet for the summary
            int summaryIndex = workbook.Worksheets.Add();
            Worksheet summarySheet = workbook.Worksheets[summaryIndex];
            summarySheet.Name = "Summary";

            // Write header row
            summarySheet.Cells["A1"].PutValue("Name");
            summarySheet.Cells["B1"].PutValue("Address");
            summarySheet.Cells["C1"].PutValue("Formula");

            int row = 1; // zero‑based index (row 2 in Excel)

            // Iterate through all named ranges
            foreach (Name namedRange in workbook.Worksheets.Names)
            {
                row++;

                // Name of the range
                summarySheet.Cells[row, 0].PutValue(namedRange.Text);

                // Address the name refers to
                summarySheet.Cells[row, 1].PutValue(namedRange.RefersTo);

                // Retrieve a formula from the first cell of the range, if any
                string formula = string.Empty;
                try
                {
                    AsposeRange range = namedRange.GetRange();
                    if (range != null && range.RowCount > 0 && range.ColumnCount > 0)
                    {
                        Cell cell = range[0, 0];
                        if (cell.IsFormula)
                        {
                            formula = cell.Formula;
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Log but ignore errors retrieving the range
                    Console.WriteLine($"Warning: could not process range '{namedRange.Text}': {ex.Message}");
                }

                summarySheet.Cells[row, 2].PutValue(formula);
            }

            // Save the workbook with the summary sheet
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Runtime error: {ex.Message}");
        }
    }
}
