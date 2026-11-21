// Title: How to set worksheet StandardWidth and then auto‑fit a specific column range using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set worksheet StandardWidth then auto fit selected columns | C# autoFitColumns after defining default column width with Aspose.Cells | override StandardWidth for a column range in an Aspose.Cells workbook | example of auto‑fitting columns B to D after setting StandardWidth in Aspose.Cells | Aspose.Cells column width default versus auto‑fit behavior in C#
// Tags: Aspose.Cells set StandardWidth C# | Aspose.Cells AutoFitColumns specific range | worksheet column width override Aspose.Cells | Excel column width default vs auto fit Aspose.Cells | C# generate Excel file with custom column widths

using Aspose.Cells;
using System;

// The sample creates a new workbook, sets the worksheet's StandardWidth to 20 characters, fills cells A1‑D5 with sample text, then auto‑fits columns B through D (indexes 1‑3), which overrides the default width for those columns, and saves the result as StandardWidth_AutoFit.xlsx.
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

            // Set the default column width (StandardWidth) for the sheet
            // The StandardWidth property is available on the Cells collection
            sheet.Cells.StandardWidth = 20; // width in characters

            // Fill some sample data into columns A‑D, rows 1‑5
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    sheet.Cells[row, col].PutValue($"Row{row + 1}Col{col + 1} sample text");
                }
            }

            // Auto‑fit columns B to D (indexes 1 to 3). This overrides StandardWidth for those columns.
            sheet.AutoFitColumns(1, 3);

            // Save the workbook
            workbook.Save("StandardWidth_AutoFit.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
