// Title: Hide rows 20‑30 in an Excel worksheet and save the workbook with Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells for .NET to hide rows 20 through 30 in a worksheet and then save the workbook as an .xlsx file. | Programmatically conceal a range of rows (20‑30) in a new Excel workbook using the Cells.HideRows method and persist the file. | Create a workbook, hide rows 20‑30 on the first sheet with Aspose.Cells, and export the result to HiddenRows.xlsx.
// Common Searches: Aspose.Cells C# hide rows 20 to 30 and save workbook | How to conceal a specific row range in Excel using Aspose.Cells .NET | C# example for hiding rows in an Aspose.Cells worksheet | Saving an Excel file with hidden rows using Aspose.Cells for .NET | Cells.HideRows method usage in Aspose.Cells C# example
// Tags: Aspose.Cells hide rows example | Cells.HideRows C# | save workbook with hidden rows Aspose | Excel row concealment Aspose.Cells .NET | programmatic row hiding Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a new workbook, hides rows 20‑30 on the first worksheet using Cells.HideRows, and saves the file as HiddenRows.xlsx.
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

            // Define the range to hide (rows 20 to 30, 1‑based indexing)
            int startRowIndex = 19; // row 20 (zero‑based)
            int endRowIndex = 29;   // row 30 (zero‑based)
            int totalRows = endRowIndex - startRowIndex + 1;

            // Hide the specified rows using the correct API
            sheet.Cells.HideRows(startRowIndex, totalRows);

            // Save the workbook with the rows concealed
            workbook.Save("HiddenRows.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
