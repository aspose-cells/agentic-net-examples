// Title: Save a merged workbook to a custom file path using Aspose.Cells in C#
// AI Prompts: Generate C# code that merges worksheets with Aspose.Cells and then saves the resulting Workbook to a user‑defined location in XLSX format. | Show a snippet that calls Workbook.Save with SaveFormat.Xlsx and a variable output path after consolidating Excel files using Aspose.Cells.
// Common Searches: how to save an Aspose.Cells merged workbook to a specific directory in C# | Aspose.Cells C# workbook.Save with dynamic file name after consolidating worksheets | C# export merged Excel workbook using Aspose.Cells to a custom path | Aspose.Cells save merged workbook to network share in C# | C# Aspose.Cells persist workbook after merging sheets
// Tags: Aspose.Cells workbook.Save Xlsx | C# save merged workbook Aspose.Cells | Aspose.Cells export merged workbook to file | custom output path Aspose.Cells | Aspose.Cells workbook persistence C#

using System;
using Aspose.Cells;

// The example defines an output file path, creates a placeholder Workbook (representing a merged workbook), and uses Aspose.Cells' Workbook.Save method with SaveFormat.Xlsx to write the file to the specified location.
class Program
{
    static void Main(string[] args)
    {
        // Specify the output file path
        string outputPath = @"C:\Temp\MergedWorkbook.xlsx";

        // Assume the merged workbook has been created/loaded earlier.
        // Here we instantiate a new workbook as a placeholder.
        Workbook workbook = new Workbook();

        // TODO: Add merging logic here (e.g., copying worksheets, consolidating data).

        // Save the merged workbook to the specified path.
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
