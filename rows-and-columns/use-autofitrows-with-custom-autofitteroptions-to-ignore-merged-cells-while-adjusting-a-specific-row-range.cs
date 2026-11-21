// Title: AutoFitRows for a specific row range when merged cells are present (C# Aspose.Cells)
// Common Searches: c# aspose.cells autofit rows range with merged cells | how to auto adjust row height for selected rows ignoring merged cells in Aspose.Cells | asp.net autofitrows specific rows while cells are merged | using AutoFitRows to fit rows 1 to 2 in an Excel workbook with Aspose.Cells
// Tags: AutoFitRows specific row range C# | Aspose.Cells ignore merged cells row height | C# adjust row height merged cells | Excel row autofit with Aspose.Cells | AutoFitterOptions row height control

using Aspose.Cells;
using System;

// The example creates a workbook, adds header and long text, merges cells A2:C2, defines a row range (rows 1‑2), calls AutoFitRows for that range using default options, and saves the file as AutoFitRows_IgnoringMergedCells.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data: put some long text in cells
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("This is a very long piece of text that would normally cause the row to expand.");
            sheet.Cells["B2"].PutValue("Another long text that could affect row height.");

            // Merge some cells in row 2 (zero‑based index 1)
            sheet.Cells.Merge(1, 0, 1, 2); // Merges A2:C2

            // Define the row range to autofit (rows 1 to 2 in Excel terms)
            int startRow = 0; // Row 1 (zero‑based)
            int endRow = 1;   // Row 2 (zero‑based)

            // Apply AutoFitRows without specifying options (default behavior)
            sheet.AutoFitRows(startRow, endRow, null);

            // Save the workbook
            workbook.Save("AutoFitRows_IgnoringMergedCells.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
