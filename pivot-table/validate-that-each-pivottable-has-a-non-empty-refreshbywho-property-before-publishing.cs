using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class ValidatePivotRefresh
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = "input.xlsx";
        string outputPath = "output_validated.xlsx";

        // Load the workbook (lifecycle rule: load)
        Workbook workbook = new Workbook(inputPath);

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all pivot tables in the current worksheet
            foreach (PivotTable pivot in sheet.PivotTables)
            {
                // Check if RefreshedByWho is empty or whitespace
                if (string.IsNullOrWhiteSpace(pivot.RefreshedByWho))
                {
                    // Refresh the pivot table to populate RefreshedByWho
                    pivot.RefreshData();
                    pivot.CalculateData();

                    // Verify again after refresh
                    if (string.IsNullOrWhiteSpace(pivot.RefreshedByWho))
                    {
                        // Validation failed – throw an exception with details
                        throw new InvalidOperationException(
                            $"PivotTable '{pivot.Name}' in worksheet '{sheet.Name}' has an empty RefreshedByWho property.");
                    }
                }
            }
        }

        // Save the workbook after successful validation (lifecycle rule: save)
        workbook.Save(outputPath);
    }
}