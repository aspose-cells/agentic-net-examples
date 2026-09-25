// Title: Recalculate formulas only on worksheets with volatile functions before exporting to PDF using Aspose.Cells for .NET
// AI Prompts: Search each worksheet for volatile Excel functions and invoke Workbook.CalculateFormula only when such functions are found, then save the workbook as a PDF. | Create a conditional formula recalculation step that runs after detecting NOW, TODAY, RAND, OFFSET, INDIRECT, INFO, or CELL functions, before converting the workbook to PDF with Aspose.Cells. | Add diagnostic logging to list worksheets containing volatile formulas, trigger on‑demand calculation, and export the workbook to PDF.
// Common Searches: Aspose.Cells recalculate formulas only on sheets with volatile functions before PDF export | C# detect volatile Excel functions in workbook and conditionally calculate formulas | how to limit formula calculation to volatile sheets when converting Excel to PDF with Aspose | conditional workbook.CalculateFormula based on volatile functions Aspose.Cells
// Tags: volatile function detection Aspose.Cells | on-demand workbook.CalculateFormula .NET | PDF conversion after formula evaluation Aspose.Cells | skip unnecessary calculations Excel to PDF | worksheet-level volatile formula check C#

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, scans each worksheet for volatile functions such as NOW, TODAY, RAND, OFFSET, INDIRECT, INFO, and CELL. If any worksheet contains these functions, it triggers a single workbook.CalculateFormula call to update formulas, then saves the workbook as a PDF, handling missing files and exceptions gracefully.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // List of common volatile functions (case‑insensitive)
            string[] volatileFunctions = new[]
            {
                "NOW()", "TODAY()", "RAND()", "RANDBETWEEN()", "OFFSET()", "INDIRECT()", "INFO()", "CELL()",
                "NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "INFO", "CELL"
            };

            bool anySheetContainsVolatile = false;

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                bool containsVolatile = false;
                Cells cells = sheet.Cells;

                // Determine the used range
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Scan cells for volatile formulas
                for (int row = 0; row <= maxRow && !containsVolatile; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula?.ToUpperInvariant() ?? string.Empty;
                            foreach (string volatileFunc in volatileFunctions)
                            {
                                if (formula.Contains(volatileFunc.ToUpperInvariant()))
                                {
                                    containsVolatile = true;
                                    break;
                                }
                            }
                        }

                        if (containsVolatile)
                            break;
                    }
                }

                if (containsVolatile)
                {
                    anySheetContainsVolatile = true;
                    // No per‑sheet calculation API; will recalculate whole workbook later
                }
            }

            // Recalculate formulas if any volatile functions were found
            if (anySheetContainsVolatile)
            {
                workbook.CalculateFormula();
            }

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
