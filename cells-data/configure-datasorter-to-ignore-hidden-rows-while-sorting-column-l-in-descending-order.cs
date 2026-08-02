using System;
using Aspose.Cells;

namespace AsposeCellsDataSorterIgnoreHidden
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate column L (zero‑based index 11) with sample data
            // Row 0 will be a header (optional)
            cells["L1"].PutValue("Value");
            for (int i = 2; i <= 10; i++) // rows 2‑10 (indices 1‑9)
            {
                cells[$"L{i}"].PutValue(10 - i); // descending values for demo
            }

            // Hide a few rows to demonstrate that they will be ignored during sorting
            cells.Rows[2].IsHidden = true; // hide row 3 (index 2)
            cells.Rows[5].IsHidden = true; // hide row 6 (index 5)

            // Configure the DataSorter
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = true;                 // first row is a header
            sorter.Key1 = 11;                         // column L (zero‑based index)
            sorter.Order1 = SortOrder.Descending;     // sort descending

            // Set AutoFitterOptions to ignore hidden rows (although this does not affect sorting,
            // it demonstrates the requested configuration)
            AutoFitterOptions autofitOptions = new AutoFitterOptions
            {
                IgnoreHidden = true,
                AutoFitMergedCells = false,
                OnlyAuto = false
            };
            // Apply the options (e.g., to autofit rows if needed)
            sheet.AutoFitRows(autofitOptions);

            // Perform the sort on the area that includes the data (rows 0‑9, column L)
            // StartRow = 0, StartColumn = 11, EndRow = 9, EndColumn = 11
            sorter.Sort(cells, 0, 11, 9, 11);

            // Save the workbook to verify the result
            workbook.Save("SortedIgnoreHidden.xlsx");
        }
    }
}