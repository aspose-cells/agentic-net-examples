// Title: AutoFit a range of rows while ignoring merged cells using AutoFitterOptions in Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to auto‑fit rows 2‑5 and sets AutoFitMergedCellsType to None. | Show how to create an AutoFitterOptions object with OnlyAuto = true and apply it to a specific row range. | Demonstrate merging cells, enabling text wrap, and then auto‑fitting rows without changing the merged cell's height.
// Common Searches: Aspose.Cells C# AutoFitRows ignore merged cells for selected rows | How to prevent merged cells from affecting row height in Aspose.Cells .NET | AutoFitterOptions OnlyAuto property usage example in Aspose.Cells | AutoFitRows specific row range with AutoFitMergedCellsType None | Adjust row height programmatically while skipping merged cells Aspose.Cells
// Tags: auto-fit rows using AutoFitterOptions Aspose.Cells | exclude merged cells from row height calculation | OnlyAuto flag for row auto‑fit | row range auto‑fit after merging cells | C# Aspose.Cells row height adjustment

using System;
using Aspose.Cells;

namespace AutoFitRowsIgnoreMergedCellsDemo
{
    // The example creates a workbook, merges cells A3:B4, enables text wrapping, configures AutoFitterOptions to ignore merged cells and affect only automatically sized rows, auto‑fits rows 2‑5 with those options, and saves the file as an XLSX document.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some rows with data
            cells["A2"].PutValue("Short text");
            cells["A3"].PutValue("This is a longer piece of text that will normally cause the row to expand.");
            cells["A4"].PutValue("Another long text that should affect row height when auto‑fitted.");
            cells["A5"].PutValue("Short again");

            // Merge cells across rows 3‑4 (zero‑based indices 2‑3) to demonstrate merged cells
            // The merged area spans 2 rows and 2 columns (A3:B4)
            cells.Merge(2, 0, 2, 2);
            cells[2, 0].PutValue("Merged cell text that would normally influence row height.");

            // Enable text wrapping for the merged cell to make the effect visible
            Style mergedStyle = cells[2, 0].GetStyle();
            mergedStyle.IsTextWrapped = true;
            cells[2, 0].SetStyle(mergedStyle);

            // Configure AutoFitterOptions to ignore merged cells
            AutoFitterOptions options = new AutoFitterOptions
            {
                // Excel's default is to ignore merged cells; explicitly set to None
                AutoFitMergedCellsType = AutoFitMergedCellsType.None,
                // Fit only rows that have not been manually sized
                OnlyAuto = true
            };

            // AutoFit rows 2 through 5 (zero‑based indices) using the options
            int startRow = 1; // Row 2 in Excel
            int endRow = 4;   // Row 5 in Excel
            worksheet.AutoFitRows(startRow, endRow, options);

            // Save the workbook to the desktop (adjust path as needed)
            string outputPath = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "AutoFitRowsIgnoreMergedCellsDemo.xlsx");
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}
