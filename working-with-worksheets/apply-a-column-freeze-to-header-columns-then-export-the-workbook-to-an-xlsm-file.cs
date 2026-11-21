// Title: Freeze the first column in an Aspose.Cells worksheet and save the workbook as a macro‑enabled XLSM file (C#)
// Common Searches: Aspose.Cells C# freeze first column of a worksheet | How to save an Aspose.Cells workbook as a macro‑enabled XLSM file | Example of using FreezePanes with Aspose.Cells to lock header columns | C# code to export a frozen worksheet to XLSM using Aspose.Cells | Create macro‑enabled Excel file with frozen panes in Aspose.Cells .NET
// Tags: freeze first column Aspose.Cells C# | export workbook to XLSM Aspose.Cells | Aspose.Cells FreezePanes method example | macro-enabled Excel saving Aspose.Cells | worksheet freeze panes C# Aspose.Cells

using Aspose.Cells;

// Demonstrates creating a new workbook, freezing the first column on the first worksheet with FreezePanes, and saving the result as a macro‑enabled XLSM file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Get the first worksheet
        var sheet = workbook.Worksheets[0];

        // Freeze the header columns (e.g., freeze the first column)
        // Parameters: row, column, totalRows, totalColumns
        // Here we freeze rows above row 0 and columns left of column 1 (first column)
        sheet.FreezePanes(0, 1, 0, 1);

        // Export the workbook to an XLSM file
        workbook.Save("output.xlsm", SaveFormat.Xlsm);
    }
}
