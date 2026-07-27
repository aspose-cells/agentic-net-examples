using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (you can customize LoadOptions if needed)
        LoadOptions loadOptions = new LoadOptions();
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Modify the value of cell B2 (row index 1, column index 1)
        worksheet.Cells["B2"].Value = "Updated long text that may need row height adjustment";

        // Auto‑fit the row that contains the modified cell (row index 1)
        worksheet.AutoFitRow(1);

        // Save the changes to a new file
        workbook.Save("output.xlsx");
    }
}

// Author: Aspose.Cells .NET example – load, edit, auto‑fit row, and save.