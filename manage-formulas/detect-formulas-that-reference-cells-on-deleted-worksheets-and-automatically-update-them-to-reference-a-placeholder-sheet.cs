using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class UpdateFormulaReferencesOnSheetDeletion
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine($"File not found: {ex.FileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException("Input workbook not found.", inputPath);

            // Load an existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Name of the worksheet that will be deleted
            string sheetToDelete = "SheetToDelete";

            // Name of the placeholder sheet that will receive the references
            string placeholderSheetName = "Placeholder";

            // Ensure the placeholder sheet exists; create it if it does not
            Worksheet placeholderSheet = null;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name.Equals(placeholderSheetName, StringComparison.OrdinalIgnoreCase))
                {
                    placeholderSheet = ws;
                    break;
                }
            }
            if (placeholderSheet == null)
            {
                placeholderSheet = workbook.Worksheets.Add(placeholderSheetName);
            }

            // Prepare the reference patterns to look for (with and without single quotes)
            string quotedRef = $"'{sheetToDelete}'!";
            string unquotedRef = $"{sheetToDelete}!";

            // Iterate through all worksheets and update formulas that reference the sheet to delete
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Skip the sheet that is about to be deleted
                if (ws.Name.Equals(sheetToDelete, StringComparison.OrdinalIgnoreCase))
                    continue;

                // Get the used range of the worksheet
                AsposeRange usedRange = ws.Cells.MaxDisplayRange;
                if (usedRange == null)
                    continue;

                // Loop through each cell in the used range
                for (int row = usedRange.FirstRow; row <= usedRange.FirstRow + usedRange.RowCount - 1; row++)
                {
                    for (int col = usedRange.FirstColumn; col <= usedRange.FirstColumn + usedRange.ColumnCount - 1; col++)
                    {
                        Cell cell = ws.Cells[row, col];
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula;
                            // Replace references to the sheet being deleted with the placeholder sheet
                            if (formula.Contains(quotedRef) || formula.Contains(unquotedRef))
                            {
                                string updatedFormula = formula
                                    .Replace(quotedRef, $"'{placeholderSheetName}'!")
                                    .Replace(unquotedRef, $"{placeholderSheetName}!");
                                cell.Formula = updatedFormula;
                            }
                        }
                    }
                }
            }

            // Find the index of the sheet to delete
            int deleteIndex = -1;
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                if (workbook.Worksheets[i].Name.Equals(sheetToDelete, StringComparison.OrdinalIgnoreCase))
                {
                    deleteIndex = i;
                    break;
                }
            }

            // Delete the worksheet if it exists
            if (deleteIndex >= 0)
            {
                workbook.Worksheets.RemoveAt(deleteIndex);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
    }
}