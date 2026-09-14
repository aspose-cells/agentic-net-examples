// Title: Replace #REF! errors with a placeholder worksheet in Excel using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that adds a worksheet named "Placeholder" if it does not exist, then scans every cell in the workbook and rewrites any formula containing "#REF!" to reference the placeholder sheet. | Create a .NET routine that loads an existing .xlsx file, ensures a fallback sheet is present, updates formulas with broken sheet references (#REF!) to point to that fallback sheet, and saves the workbook.
// Common Searches: Aspose.Cells C# replace #REF! errors after deleting a sheet | How to automatically fix broken worksheet references in an Excel file using Aspose.Cells | C# code to change formulas containing #REF! to point to a placeholder sheet | Detect and correct invalid sheet references in .xlsx with Aspose.Cells .NET
// Tags: Aspose.Cells replace broken sheet references | C# update #REF! formulas in Excel | placeholder worksheet for invalid references | scan workbook cells for formula errors Aspose.Cells | auto-correct Excel formulas after sheet deletion .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program loads an Excel workbook (or creates a new one), guarantees a worksheet named "Placeholder" exists, iterates through all used cells of each non‑placeholder sheet, replaces any formula containing "#REF!" with a reference to the placeholder sheet, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";
                const string placeholderSheetName = "Placeholder";

                // Load existing workbook if the file exists; otherwise create a new workbook
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Ensure a placeholder worksheet exists
                bool placeholderExists = false;
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Name.Equals(placeholderSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        placeholderExists = true;
                        break;
                    }
                }
                if (!placeholderExists)
                {
                    workbook.Worksheets.Add(placeholderSheetName);
                }

                // Iterate through worksheets (skip the placeholder) and fix formulas containing #REF!
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Name.Equals(placeholderSheetName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    // Get the used range of the sheet
                    Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                    if (usedRange == null)
                        continue;

                    foreach (Cell cell in usedRange)
                    {
                        // Check if the cell contains a formula
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula;
                            if (!string.IsNullOrEmpty(formula) && formula.Contains("#REF!"))
                            {
                                // Replace broken reference with placeholder sheet name
                                string updatedFormula = formula.Replace("#REF!", placeholderSheetName + "!");
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
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
