using System;
using System.IO;
using Aspose.Cells;

namespace AutoFitRowsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate rows 2 to 5 with sample data
            for (int row = 2; row <= 5; row++)
            {
                cells[row, 0].PutValue($"Row {row} - This is a long text that should cause the row height to increase when auto‑fitted.");
                cells[row, 1].PutValue($"Additional column data for row {row}.");
            }

            // Merge some cells across rows 2‑3 and columns 0‑1 to demonstrate merged cells
            // The merged area will be ignored during auto‑fit because of the option we set later
            cells.Merge(2, 0, 2, 2);
            cells[2, 0].PutValue("Merged cell text that would normally affect row height.");

            // Enable text wrapping for the merged cell (optional, just for realism)
            Style mergedStyle = cells[2, 0].GetStyle();
            mergedStyle.IsTextWrapped = true;
            cells[2, 0].SetStyle(mergedStyle);

            // Configure AutoFitterOptions to ignore merged cells (AutoFitMergedCellsType.None)
            AutoFitterOptions options = new AutoFitterOptions
            {
                AutoFitMergedCellsType = AutoFitMergedCellsType.None,
                OnlyAuto = true   // Fit only rows that have not been manually sized
            };

            // Auto‑fit rows 2 through 5 using the options (rule: AutoFitRows(int, int, AutoFitterOptions))
            sheet.AutoFitRows(2, 5, options);

            // Save the workbook (lifecycle rule: save)
            string outputPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "AutoFitRows_IgnoringMergedCells.xlsx");
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}