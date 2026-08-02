// Title: AutoFit rows while ignoring merged cells with AutoFitterOptions in Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, merge cells, enable text wrapping, and use AutoFitterOptions (AutoFitMergedCellsType.None, OnlyAuto = true) to auto‑fit the height of rows 0‑3 without affecting merged cells or rows that already have a custom height, then save the file.
// Keywords: Aspose.Cells | AutoFitRows | AutoFitterOptions | AutoFitMergedCellsType.None | OnlyAuto | ignore merged cells | C# | .NET | row height | merged cells | range autofit
// Common Searches: Aspose.Cells auto fit rows ignore merged cells | AutoFitterOptions OnlyAuto example C# | How to skip merged cells when auto fitting rows Aspose.Cells | AutoFitRows specific range .NET | Adjust row height without affecting merged cells Aspose.Cells
// Developer Intent: Automatically adjust the height of selected rows while excluding merged cells and preserving any rows with manually set heights.
// Use Cases: Exporting tabular reports where header rows are merged and must retain a fixed height. | Generating spreadsheets where only data rows need dynamic height, leaving merged title rows unchanged. | Applying row auto‑fit to a subset of rows in a sheet that contains merged cells, without disturbing the layout.
// AI Prompts: Generate C# code that auto‑fits rows 10‑15 in an Aspose.Cells worksheet while ignoring merged cells and keeping manually sized rows unchanged. | Show how to configure AutoFitterOptions to skip merged cells and auto‑fit only default‑height rows in Aspose.Cells. | Explain the impact of AutoFitMergedCellsType.None and OnlyAuto on the AutoFitRows method.

using System;
using Aspose.Cells;

// Shows how to create a workbook, merge cells, enable text wrapping, and use AutoFitterOptions (AutoFitMergedCellsType.None, OnlyAuto = true) to auto‑fit the height of rows 0‑3 without affecting merged cells or rows that already have a custom height, then save the file.
class AutoFitRowsIgnoreMergedDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate some sample data
        cells["A1"].PutValue("Short text");
        cells["A2"].PutValue("This is a longer text that should cause the row height to increase when auto‑fitted.");
        cells["A3"].PutValue("Merged cell text that would normally affect row height across multiple rows.");
        cells["B3"].PutValue(""); // part of the merged area

        // Merge cells A3:B3 (row index 2) to create a merged cell
        cells.Merge(2, 0, 1, 2);
        // Enable text wrapping for the merged cell
        Style mergedStyle = cells["A3"].GetStyle();
        mergedStyle.IsTextWrapped = true;
        cells["A3"].SetStyle(mergedStyle);

        // Configure AutoFitterOptions to ignore merged cells
        AutoFitterOptions options = new AutoFitterOptions
        {
            AutoFitMergedCellsType = AutoFitMergedCellsType.None, // ignore merged cells during autofit
            OnlyAuto = true // only autofit rows that don't have a custom height
        };

        // Auto‑fit rows 0 to 3 (first four rows) using the options
        sheet.AutoFitRows(0, 3, options);

        // Save the workbook
        workbook.Save("AutoFitRowsIgnoreMerged.xlsx");
    }
}
