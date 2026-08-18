// Title: C# – AutoFit specific rows while ignoring merged cells with AutoFitterOptions (Aspose.Cells)
// Description: Creates a workbook, writes wrapped text to cells A1‑A5, merges A5:B5, configures AutoFitterOptions to skip merged cells and affect only rows with default height, then auto‑fits rows 0‑4 and saves the file.
// Keywords: Aspose.Cells AutoFitRows | AutoFitterOptions | ignore merged cells | OnlyAuto property | C# row height auto‑fit | specific row range
// Common Searches: Aspose.Cells auto‑fit rows ignore merged cells | AutoFitRows with AutoFitterOptions C# | fit rows 0‑4 only default height Aspose | skip merged cells when auto‑sizing rows
// Developer Intent: Adjust the height of rows 0‑4 without letting merged cells influence the calculation.
// Use Cases: Generate reports where header merges must not change row height. | Auto‑size wrapped‑text rows while preserving manually set row heights. | Apply row auto‑fit after merging cells in dynamic spreadsheets.
// AI Prompts: Write C# code that auto‑fits rows 2‑6 in an Aspose.Cells workbook while ignoring merged cells and keeping custom row heights unchanged. | Explain the impact of AutoFitterOptions.AutoFitMergedCellsType and OnlyAuto on AutoFitRows behavior. | Show how to auto‑fit a row range after merging cells without altering the merged cell's height.

using System;
using Aspose.Cells;

namespace AutoFitRowsExample
{
    // Creates a workbook, writes wrapped text to cells A1‑A5, merges A5:B5, configures AutoFitterOptions to skip merged cells and affect only rows with default height, then auto‑fits rows 0‑4 and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some rows with data
            cells["A1"].PutValue("Short text");
            cells["A2"].PutValue("This is a longer piece of text that will normally increase row height.");
            cells["A3"].PutValue("Another long text that spans multiple lines when wrapped.\nLine 2.\nLine 3.");
            cells["A4"].PutValue("Normal text");
            cells["A5"].PutValue("Merged cell content that would affect row heights if not ignored.");

            // Apply text wrapping to demonstrate row height changes
            Style wrapStyle = cells["A2"].GetStyle();
            wrapStyle.IsTextWrapped = true;
            cells["A2"].SetStyle(wrapStyle);
            cells["A3"].SetStyle(wrapStyle);
            cells["A5"].SetStyle(wrapStyle);

            // Merge cells A5:B5 (rows 4, columns 0-1) – this merged cell will be ignored during autofit
            cells.Merge(4, 0, 1, 2);

            // Configure AutoFitterOptions to ignore merged cells
            AutoFitterOptions options = new AutoFitterOptions
            {
                AutoFitMergedCellsType = AutoFitMergedCellsType.None, // ignore merged cells
                OnlyAuto = true   // fit only rows without custom height
            };

            // AutoFit rows 0 through 4 using the options (lifecycle rule: method overload)
            sheet.AutoFitRows(0, 4, options);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("AutoFitRows_IgnoringMergedCells.xlsx");
        }
    }
}
