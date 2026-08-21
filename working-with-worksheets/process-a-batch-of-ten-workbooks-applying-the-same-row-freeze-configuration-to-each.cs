// Title: Batch freeze panes in multiple Excel workbooks with Aspose.Cells for .NET (C#)
// Description: A C# console app that loads ten Excel files, applies the same FreezePanes setting (cell C3, 3 rows × 3 columns) to the first worksheet of each workbook, and saves the updated files to a "ProcessedWorkbooks" directory. Ideal for automating consistent view layouts across many spreadsheets.
// Keywords: Aspose.Cells | C# | .NET | freeze panes | batch processing | multiple workbooks | Excel automation | FreezePanes method | C3 cell | rows and columns | programmatic Excel
// Common Searches: Aspose.Cells batch freeze panes C# | How to apply same freeze pane to many Excel files | Freeze rows and columns in multiple workbooks using .NET | Programmatic FreezePanes for a list of Excel files | Automate Excel view settings with Aspose.Cells
// Developer Intent: Programmatically set an identical freeze‑pane configuration on a collection of Excel workbooks in a single run.
// Use Cases: Standardize header visibility for a series of monthly reports before distribution. | Prepare template workbooks with frozen panes so end users get a consistent navigation experience. | Pre‑process uploaded Excel files on a server to enforce a uniform layout for downstream analytics.
// AI Prompts: Create C# code using Aspose.Cells that freezes the first 4 rows and 2 columns in every worksheet of all Excel files in a specified folder. | Write a script that iterates over a list of workbook paths, applies a freeze pane at cell B2, and saves the modified files to a separate output directory. | Show an example of batch processing Excel workbooks to set a freeze pane at D5 while preserving existing formatting, using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

// A C# console app that loads ten Excel files, applies the same FreezePanes setting (cell C3, 3 rows × 3 columns) to the first worksheet of each workbook, and saves the updated files to a "ProcessedWorkbooks" directory. Ideal for automating consistent view layouts across many spreadsheets.
class BatchFreezePanes
{
    static void Main()
    {
        // Define the list of workbook file paths to process (10 files)
        string[] inputFiles = new string[]
        {
            "Workbook1.xlsx",
            "Workbook2.xlsx",
            "Workbook3.xlsx",
            "Workbook4.xlsx",
            "Workbook5.xlsx",
            "Workbook6.xlsx",
            "Workbook7.xlsx",
            "Workbook8.xlsx",
            "Workbook9.xlsx",
            "Workbook10.xlsx"
        };

        // Ensure the output directory exists
        string outputDir = "ProcessedWorkbooks";
        Directory.CreateDirectory(outputDir);

        // Freeze configuration: freeze at cell C3 with 3 rows and 3 columns frozen
        string freezeCell = "C3";
        int frozenRows = 3;
        int frozenColumns = 3;

        foreach (string inputPath in inputFiles)
        {
            // Load the workbook from file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Apply the freeze panes setting
            worksheet.FreezePanes(freezeCell, frozenRows, frozenColumns);

            // Build the output file path
            string outputPath = Path.Combine(outputDir, Path.GetFileName(inputPath));

            // Save the modified workbook
            workbook.Save(outputPath);
        }

        Console.WriteLine("Batch processing completed. Modified workbooks are saved in '" + outputDir + "'.");
    }
}
