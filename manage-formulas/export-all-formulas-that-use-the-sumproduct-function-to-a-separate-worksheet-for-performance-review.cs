// Title: Export all SUMPRODUCT formulas to a new worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example to list cells containing SUMPRODUCT formula | How to create a summary sheet of SUMPRODUCT functions with Aspose.Cells | Extract Excel formulas that use SUMPRODUCT into a separate worksheet in .NET | Iterate through workbook cells and filter by SUMPRODUCT using Aspose.Cells | Save workbook after adding a sheet for SUMPRODUCT formula export in C#
// Tags: export SUMPRODUCT formulas Aspose.Cells | iterate workbook cells C# Aspose.Cells | filter Excel formulas by function name .NET | add summary worksheet for specific formulas Aspose.Cells | save modified workbook after formula extraction C# | create header row in new worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace SumProductExport
{
    // The C# program loads an existing Excel workbook, adds a new worksheet named "SUMPRODUCT_Formulas", writes a header row, scans every used cell in all other worksheets, detects cells whose formula contains the SUMPRODUCT function (case‑insensitive), and records the source sheet name, cell address, and formula in the new sheet before saving the workbook to the specified output file.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths (replace with actual paths)
            string inputFile = "input.xlsx";
            string outputFile = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputFile);

                // Add a new worksheet for SUMPRODUCT formulas
                int exportSheetIndex = workbook.Worksheets.Add();
                Worksheet exportSheet = workbook.Worksheets[exportSheetIndex];
                exportSheet.Name = "SUMPRODUCT_Formulas";

                // Write header row
                exportSheet.Cells[0, 0].PutValue("Source Sheet");
                exportSheet.Cells[0, 1].PutValue("Cell Address");
                exportSheet.Cells[0, 2].PutValue("Formula");

                int exportRow = 1; // Start after header

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Skip the export sheet itself
                    if (sheet.Name == exportSheet.Name) continue;

                    // Determine the used range
                    int maxRow = sheet.Cells.MaxDataRow;
                    int maxCol = sheet.Cells.MaxDataColumn;

                    // Scan each cell in the used range
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];
                            if (cell.IsFormula && !string.IsNullOrEmpty(cell.Formula) &&
                                cell.Formula.IndexOf("SUMPRODUCT", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // Export details to the new worksheet
                                exportSheet.Cells[exportRow, 0].PutValue(sheet.Name);
                                exportSheet.Cells[exportRow, 1].PutValue(cell.Name);
                                exportSheet.Cells[exportRow, 2].PutValue(cell.Formula);
                                exportRow++;
                            }
                        }
                    }
                }

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputFile);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the new worksheet
                workbook.Save(outputFile);
                Console.WriteLine($"Export completed. File saved to {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
