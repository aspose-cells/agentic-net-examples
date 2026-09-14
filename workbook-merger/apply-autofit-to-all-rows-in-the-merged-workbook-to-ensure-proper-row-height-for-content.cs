// Title: AutoFit row heights for all worksheets in a merged Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens a merged Excel file, sets AutoFitterOptions to expand rows that contain merged cells, and calls AutoFitRows on every worksheet. | Update an existing workbook merger sample to include AutoFitRows with OnlyAuto=true so rows with custom heights are left untouched. | Create a script that processes all sheets, applies the auto‑fit operation, and saves the result to a new file name.
// Common Searches: asp.net how to autofit rows in merged cells after merging workbooks | c# Aspose.Cells AutoFitRows eachline option example | adjust row height for merged cells across all worksheets using Aspose.Cells | auto fit rows only when height not set Aspose.Cells .NET
// Tags: auto-fit rows Aspose.Cells .NET | AutoFitterOptions eachline merged cells | apply AutoFitRows to all worksheets | row height adjustment after workbook merge | C# Aspose.Cells row height auto-fit

using System;
using Aspose.Cells;

// Loads a merged workbook, configures AutoFitterOptions to expand each line of merged cells while only auto‑fitting rows without custom heights, applies AutoFitRows to every worksheet, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        // Load the merged workbook (replace with actual file path)
        string inputPath = "mergedWorkbook.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Configure AutoFitterOptions to handle merged cells correctly
        AutoFitterOptions options = new AutoFitterOptions
        {
            // Expand the height of each row that participates in a merged cell
            AutoFitMergedCellsType = AutoFitMergedCellsType.EachLine,
            // Fit only rows that have not been given a custom height
            OnlyAuto = true
        };

        // Apply AutoFitRows to every worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.AutoFitRows(options);
        }

        // Save the updated workbook
        string outputPath = "mergedWorkbook_AutoFitRows.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
