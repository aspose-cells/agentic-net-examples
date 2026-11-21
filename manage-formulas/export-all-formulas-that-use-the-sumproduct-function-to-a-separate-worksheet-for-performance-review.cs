// Title: Export all SUMPRODUCT formulas to a new worksheet with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells code to copy cells that use SUMPRODUCT into another sheet | How to list Excel formulas containing SUMPRODUCT using Aspose.Cells library | Extract SUMPRODUCT function formulas from a workbook and save to a separate worksheet in .NET | Aspose.Cells iterate through workbook cells and filter by SUMPRODUCT formula | Create a summary sheet of SUMPRODUCT formulas with source address using Aspose.Cells C#
// Tags: export SUMPRODUCT formulas Aspose.Cells | list cells by function name C# | filter workbook formulas .NET | write formula details to new worksheet | extract Excel formulas using Aspose.Cells | create summary sheet for specific formulas

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook (or creates a new one if missing), adds a worksheet named "SUMPRODUCT_Formulas", writes a header row, then iterates through every worksheet (excluding the export sheet) to find cells whose formulas contain the SUMPRODUCT function. For each match it records the source sheet name, cell address, and the full formula on the new sheet, and finally saves the workbook with the added summary worksheet.
class ExportSumProductFormulas
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Ensure the input file exists; if not, create an empty workbook to avoid FileNotFoundException
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Add a new worksheet to hold the SUMPRODUCT formulas
            int exportSheetIndex = workbook.Worksheets.Add();
            Worksheet exportSheet = workbook.Worksheets[exportSheetIndex];
            exportSheet.Name = "SUMPRODUCT_Formulas";

            // Write header row
            exportSheet.Cells[0, 0].PutValue("Source Sheet");
            exportSheet.Cells[0, 1].PutValue("Cell Address");
            exportSheet.Cells[0, 2].PutValue("Formula");

            int exportRow = 1; // start after header

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Skip the export sheet itself to avoid self‑reference
                if (sheet.Name == exportSheet.Name) continue;

                // Get the used range of the current sheet
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                // Iterate through each cell in the used range
                foreach (Cell cell in usedRange)
                {
                    // Check if the cell contains a formula
                    if (cell.IsFormula)
                    {
                        // Look for the SUMPRODUCT function (case‑insensitive)
                        if (cell.Formula.IndexOf("SUMPRODUCT", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Write details to the export sheet
                            exportSheet.Cells[exportRow, 0].PutValue(sheet.Name);
                            exportSheet.Cells[exportRow, 1].PutValue(cell.Name); // e.g., "B5"
                            exportSheet.Cells[exportRow, 2].PutValue(cell.Formula);
                            exportRow++;
                        }
                    }
                }
            }

            // Save the workbook with the new worksheet
            workbook.Save(outputPath);
            Console.WriteLine($"Export completed. Output saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
