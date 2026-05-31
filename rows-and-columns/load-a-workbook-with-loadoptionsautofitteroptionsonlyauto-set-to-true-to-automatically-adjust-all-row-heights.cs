using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Create LoadOptions and configure AutoFitterOptions to only auto‑fit rows that are not custom‑sized
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.AutoFitterOptions = new AutoFitterOptions();
        loadOptions.AutoFitterOptions.OnlyAuto = true;

        // Load the workbook with the specified options; rows will be auto‑fitted during loading
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Example: display the height of the first row after auto‑fit
        double firstRowHeight = workbook.Worksheets[0].Cells.GetRowHeight(0);
        Console.WriteLine($"First row height after auto‑fit: {firstRowHeight}");

        // Save the workbook to a new file
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}