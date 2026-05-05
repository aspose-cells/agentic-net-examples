using System;
using Aspose.Cells;

namespace AsposeCellsGroupDuringLoad
{
    class Program
    {
        static void Main()
        {
            // Define load options (you can customize as needed)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook with the specified options
            Workbook workbook = new Workbook("InputFile.xlsx", loadOptions);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Apply grouping to rows 2‑6 (zero‑based index 1‑5) and keep them visible
            cells.GroupRows(1, 5, false);

            // Optionally control outline settings (e.g., place summary row above detail)
            worksheet.Outline.SummaryRowBelow = false;

            // Save the modified workbook
            workbook.Save("GroupedOutput.xlsx");
        }
    }
}