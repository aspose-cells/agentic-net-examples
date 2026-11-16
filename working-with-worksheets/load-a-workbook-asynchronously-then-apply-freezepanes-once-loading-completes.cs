// Title: Load an Excel workbook asynchronously and set FreezePanes on the first worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load workbook asynchronously and then freeze first row and column | how to apply FreezePanes after async workbook load in .NET | example of using Task.Run to open Excel file with Aspose.Cells and set freeze panes | save workbook after applying FreezePanes on first sheet in C# Aspose.Cells
// Tags: asynchronous workbook loading Aspose.Cells | apply FreezePanes first worksheet Aspose.Cells | freeze first row and column Excel Aspose.Cells | save modified workbook after async load Aspose.Cells | C# worksheet manipulation async Aspose.Cells

using Aspose.Cells;
using System;
using System.Threading.Tasks;

// The sample loads 'input.xlsx' asynchronously into an Aspose.Cells Workbook, accesses the first worksheet, applies FreezePanes at row 1 and column 1 to lock the top row and left column, and then saves the result as 'output.xlsx'.
class Program
{
    static async Task Main()
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";

        // Load the workbook asynchronously
        Workbook workbook = await Task.Run(() => new Workbook(inputPath));

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Apply FreezePanes:
        // Freeze the first row and the first column (split at row 1, column 1)
        // Parameters: splitRow, splitColumn, leftColumn, topRow
        sheet.FreezePanes(1, 1, 1, 1);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
