// Title: C# – Detect #REF! Formulas After Deleting a Worksheet and Redirect Them to a Placeholder Sheet with Aspose.Cells
// Description: Creates a workbook, adds a placeholder sheet, writes formulas that point to a sheet scheduled for removal, deletes that sheet, scans all cells for "#REF!" errors, replaces the broken reference with the placeholder sheet name, and saves the corrected file.
// Keywords: Aspose.Cells C# detect #REF! | update broken worksheet references | replace deleted sheet formulas | placeholder sheet Aspose.Cells | scan workbook for #REF! errors | auto‑fix Excel references .NET
// Common Searches: Aspose.Cells replace #REF! after deleting sheet | C# find broken formulas in Excel workbook | redirect deleted worksheet references to placeholder | how to fix #REF! errors with Aspose.Cells | auto‑update formulas when sheet is removed
// Developer Intent: Locate formulas that become #REF! after a worksheet is removed and automatically point them to a designated placeholder sheet.
// Use Cases: Programmatically delete a worksheet and ensure no #REF! errors remain. | Maintain workbook integrity by redirecting all broken references to a placeholder sheet. | Generate clean Excel files for downstream processing or reporting.
// AI Prompts: Generate C# code using Aspose.Cells that scans a Workbook for #REF! formulas after a sheet deletion and changes them to reference a sheet named "Placeholder". | Write a reusable method that accepts a Workbook and a placeholder sheet name, then updates any formulas broken by removed worksheets. | Explain the steps to safely delete a worksheet and automatically redirect its formulas to a placeholder sheet with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, adds a placeholder sheet, writes formulas that point to a sheet scheduled for removal, deletes that sheet, scans all cells for "#REF!" errors, replaces the broken reference with the placeholder sheet name, and saves the corrected file.
class DetectAndUpdateDeletedSheetFormulas
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a placeholder sheet that will receive the broken references
            Worksheet placeholderSheet = workbook.Worksheets.Add("Placeholder");

            // Add a sheet that will be deleted later
            Worksheet sheetToDelete = workbook.Worksheets.Add("SheetToDelete");

            // Add a main sheet with formulas referencing the sheet that will be deleted
            Worksheet mainSheet = workbook.Worksheets[0];
            mainSheet.Name = "Main";
            mainSheet.Cells["A1"].Formula = "=SheetToDelete!B1";
            mainSheet.Cells["A2"].Formula = "=SUM(SheetToDelete!B1:B5)";

            // Populate some data in the sheet that will be deleted
            sheetToDelete.Cells["B1"].PutValue(10);
            sheetToDelete.Cells["B2"].PutValue(20);
            sheetToDelete.Cells["B3"].PutValue(30);
            sheetToDelete.Cells["B4"].PutValue(40);
            sheetToDelete.Cells["B5"].PutValue(50);

            // Optional: calculate formulas before deletion
            workbook.CalculateFormula();

            // Delete the sheet that contains the referenced cells
            int deleteIndex = workbook.Worksheets.IndexOf(sheetToDelete);
            if (deleteIndex != -1)
            {
                workbook.Worksheets.RemoveAt(deleteIndex);
            }

            // After deletion, formulas that referenced the removed sheet become "#REF!"
            // Iterate through all worksheets and cells to replace "#REF!" with the placeholder sheet name
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Cells cells = ws.Cells;
                foreach (Cell cell in cells)
                {
                    if (cell.IsFormula && cell.Formula.Contains("#REF!"))
                    {
                        // Update the broken reference to point to the placeholder sheet
                        cell.Formula = cell.Formula.Replace("#REF!", placeholderSheet.Name);
                    }
                }
            }

            // Define output file path
            string outputPath = "UpdatedFormulas.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
