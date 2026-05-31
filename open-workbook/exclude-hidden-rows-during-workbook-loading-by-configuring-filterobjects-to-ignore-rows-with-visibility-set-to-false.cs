using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Prepare load options to enable auto‑filter processing while loading
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.AutoFilter = true; // activates filtering of rows based on AutoFilter criteria

        // Load the workbook with the specified options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Access the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Refresh the AutoFilter – this applies the filter and returns the indexes of hidden rows
        int[] hiddenRowIndexes = sheet.AutoFilter.Refresh();

        // If there are hidden rows, remove them so they are excluded from the loaded data
        if (hiddenRowIndexes != null && hiddenRowIndexes.Length > 0)
        {
            // Remove rows from bottom to top to keep subsequent indexes valid
            for (int i = hiddenRowIndexes.Length - 1; i >= 0; i--)
            {
                sheet.Cells.Rows.RemoveAt(hiddenRowIndexes[i]);
            }
        }

        // Save the workbook after hidden rows have been excluded
        workbook.Save("output.xlsx");
    }
}