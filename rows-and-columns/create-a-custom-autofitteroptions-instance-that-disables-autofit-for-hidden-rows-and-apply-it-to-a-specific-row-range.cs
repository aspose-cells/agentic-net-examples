using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data to rows 0‑4
        for (int i = 0; i < 5; i++)
        {
            worksheet.Cells[i, 0].PutValue($"This is a long text in row {i} that may require autofit.");
        }

        // Hide row 2 to demonstrate the effect of IgnoreHidden
        worksheet.Cells.Rows[2].IsHidden = true;

        // Create AutoFitterOptions with IgnoreHidden set to false
        // This disables auto‑fit for hidden rows (row 2 will keep its original height)
        AutoFitterOptions options = new AutoFitterOptions
        {
            IgnoreHidden = false
        };

        // Apply autofit to rows 0 through 4 using the custom options
        worksheet.AutoFitRows(0, 4, options);

        // Save the workbook
        workbook.Save("AutoFitRowsExample.xlsx");
    }
}

// Author: Aspose.Cells .NET example