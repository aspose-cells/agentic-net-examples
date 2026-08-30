// Title: Load an Excel workbook with OnlyAuto auto‑fit rows, iterate all worksheets, and verify each row’s height using Aspose.Cells for .NET
// AI Prompts: Load a .xlsx file with LoadOptions.AutoFitterOptions.OnlyAuto = true, then print the height and IsHeightMatched flag for every row in each worksheet. | Write a helper method that accepts a file path, loads the workbook with OnlyAuto auto‑fit enabled, and returns a dictionary mapping worksheet names to a list of (row index, height, IsHeightMatched) tuples. | Modify the sample to persist any changes by saving the workbook after the row‑height verification while keeping OnlyAuto behavior unchanged.
// Common Searches: Aspose.Cells OnlyAuto option how to load workbook and get row heights in C# | C# iterate worksheets and check IsHeightMatched after loading Excel with auto‑fit rows only | How to verify that rows were auto‑fitted when loading an XLSX file using Aspose.Cells | Retrieve row height values from a workbook loaded with AutoFitterOptions.OnlyAuto in .NET
// Tags: AutoFitterOptions OnlyAuto configuration | Aspose.Cells extract row height C# | loop through worksheets row metrics | IsHeightMatched flag evaluation | auto‑fit rows only Excel import

using System;
using Aspose.Cells;

// The example sets LoadOptions.AutoFitterOptions.OnlyAuto to true, loads an Excel file, iterates each worksheet, and prints the height and IsHeightMatched flag for every row up to the last data row, then optionally saves the workbook.
class Program
{
    static void Main()
    {
        // Configure load options to enable OnlyAuto (auto‑fit only rows without custom height)
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.AutoFitterOptions = new AutoFitterOptions
        {
            OnlyAuto = true
        };

        // Load the workbook with the specified options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Iterate through each worksheet and output row heights
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Worksheet: {sheet.Name}");

            // Determine the last row that contains data
            int lastDataRow = sheet.Cells.MaxDataRow;

            for (int rowIndex = 0; rowIndex <= lastDataRow; rowIndex++)
            {
                double height = sheet.Cells.GetRowHeight(rowIndex);
                bool isHeightMatched = sheet.Cells.Rows[rowIndex].IsHeightMatched;
                Console.WriteLine($"Row {rowIndex} height: {height} (IsHeightMatched = {isHeightMatched})");
            }
        }

        // Save the workbook (optional, to persist any changes)
        workbook.Save("output.xlsx");
    }
}
