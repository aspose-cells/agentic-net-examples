// Title: C# code to detect #REF! formulas caused by deleted worksheets and replace them with a placeholder sheet using Aspose.Cells
// Common Searches: Aspose.Cells C# replace #REF! errors after deleting a worksheet | how to fix broken formulas referencing removed sheets in .NET Excel workbook | detect and update formulas containing #REF! using Aspose.Cells for .xlsx files | C# iterate cells to change #REF! to placeholder sheet in Excel workbook | automate placeholder sheet creation for broken references with Aspose.Cells
// Tags: Aspose.Cells replace #REF! formulas | C# placeholder worksheet for broken references | detect deleted sheet references in Excel | iterate cells to fix formulas Aspose.Cells | update Excel #REF! errors .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook (or creates a new one), ensures a worksheet named "Placeholder" exists, scans every non‑placeholder sheet for formulas containing "#REF!", replaces the broken reference with "Placeholder!A1", and saves the updated workbook.
class UpdateFormulasAfterSheetDeletion
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string placeholderName = "Placeholder";

        Workbook workbook;

        // Load workbook if it exists; otherwise create a new one
        if (File.Exists(inputPath))
        {
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (CellsException ex)
            {
                // Handle password‑protected or corrupted files gracefully
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                Console.WriteLine("Creating a new workbook instead.");
                workbook = new Workbook();
            }
        }
        else
        {
            Console.WriteLine("Input file not found. Creating a new workbook.");
            workbook = new Workbook();
        }

        try
        {
            // Ensure a placeholder sheet exists (create if missing)
            Worksheet placeholderSheet = null;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name.Equals(placeholderName, StringComparison.OrdinalIgnoreCase))
                {
                    placeholderSheet = ws;
                    break;
                }
            }

            if (placeholderSheet == null)
            {
                placeholderSheet = workbook.Worksheets[workbook.Worksheets.Add()];
                placeholderSheet.Name = placeholderName;
                placeholderSheet.Cells["A1"].PutValue("Placeholder");
            }

            // Iterate through all worksheets and cells to fix formulas containing #REF!
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name.Equals(placeholderName, StringComparison.OrdinalIgnoreCase))
                    continue; // skip placeholder sheet

                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell.IsFormula && !string.IsNullOrEmpty(cell.Formula) && cell.Formula.Contains("#REF!"))
                        {
                            // Replace broken reference with placeholder cell reference
                            string updatedFormula = cell.Formula.Replace("#REF!", $"{placeholderName}!A1");
                            cell.Formula = updatedFormula;
                        }
                    }
                }
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
