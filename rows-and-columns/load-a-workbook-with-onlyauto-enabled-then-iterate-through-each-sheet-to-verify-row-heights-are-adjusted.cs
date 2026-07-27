using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load workbook with OnlyAuto enabled (auto‑fit only rows without custom height)
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.AutoFitterOptions = new AutoFitterOptions();
        loadOptions.AutoFitterOptions.OnlyAuto = true;
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Iterate through each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Auto‑fit rows, respecting the OnlyAuto flag
            sheet.AutoFitRows(true);

            // Verify row heights
            int lastRow = sheet.Cells.MaxDataRow;
            for (int i = 0; i <= lastRow; i++)
            {
                Row row = sheet.Cells.Rows[i];
                Console.WriteLine($"Sheet: {sheet.Name}, Row: {i}, Height: {row.Height}, IsHeightMatched: {row.IsHeightMatched}");
            }
        }

        // Save the workbook after verification
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}

// Author: Aspose.Cells .NET example – loads with OnlyAuto, auto‑fits rows, and checks heights.