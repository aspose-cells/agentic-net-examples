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
        for (int i = 0; i <= 4; i++)
        {
            worksheet.Cells[i, 0].PutValue($"Row {i} contains a long piece of text that may affect row height.");
        }

        // Hide row 2 (index 2)
        worksheet.Cells.Rows[2].IsHidden = true;

        // Create AutoFitterOptions that ignores hidden rows/columns
        AutoFitterOptions options = new AutoFitterOptions
        {
            IgnoreHidden = true   // disables auto‑fit for hidden rows
        };

        // Apply autofit to rows 0 through 4 using the custom options
        worksheet.AutoFitRows(0, 4, options);

        // Save the workbook
        workbook.Save("AutoFitRowsIgnoreHidden.xlsx");
    }
}