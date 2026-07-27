using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Author: Aspose.Cells .NET example
    class AutoFitRowsOnLoad
    {
        static void Main()
        {
            // Create AutoFitterOptions to control row auto‑fitting during load
            AutoFitterOptions fitterOptions = new AutoFitterOptions
            {
                // False – auto‑fit all rows (including those with custom height)
                OnlyAuto = false,
                // Optional: also auto‑fit merged cells if needed
                AutoFitMergedCells = true
            };

            // Assign the options to LoadOptions
            LoadOptions loadOptions = new LoadOptions
            {
                AutoFitterOptions = fitterOptions
            };

            // Load the workbook with the specified options; rows will be auto‑fitted
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Example: retrieve a row height after auto‑fit
            double firstRowHeight = workbook.Worksheets[0].Cells.GetRowHeight(0);
            Console.WriteLine($"First row height after auto‑fit: {firstRowHeight}");

            // Save the workbook (preserving the adjusted row heights)
            workbook.Save("output.xlsx");
        }
    }
}