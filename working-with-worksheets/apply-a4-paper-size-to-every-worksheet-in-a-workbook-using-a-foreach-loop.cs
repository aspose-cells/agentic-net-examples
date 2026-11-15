// Title: Set A4 paper size for every worksheet in an Aspose.Cells workbook using a C# foreach loop
// Common Searches: C# foreach loop to set A4 paper size on all worksheets with Aspose.Cells | How to change page setup paper size for each sheet in an Aspose.Cells workbook | Apply default A4 page size to new worksheets using Aspose.Cells API | Iterate through worksheets and set PaperSizeType.PaperA4 in C# | Aspose.Cells set page setup paper size for multiple worksheets programmatically
// Tags: Aspose.Cells set worksheet paper size | C# foreach iterate worksheets page setup | apply A4 paper size Aspose.Cells workbook | automate page setup for Excel sheets C# | paper size configuration Aspose.Cells API

using Aspose.Cells;
using System;

// The example creates a Workbook, adds two uniquely named worksheets, then uses a foreach loop to set each worksheet’s PageSetup.PaperSize to PaperSizeType.PaperA4, and finally saves the file as Output.xlsx. A helper method checks for existing worksheet names.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (contains a default worksheet)
            Workbook workbook = new Workbook();

            // Add additional worksheets with unique names, avoiding duplicates
            if (!WorksheetExists(workbook, "Sheet1"))
                workbook.Worksheets.Add("Sheet1");
            if (!WorksheetExists(workbook, "Sheet2"))
                workbook.Worksheets.Add("Sheet2");

            // Apply A4 paper size to every worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
            }

            // Save the workbook
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Checks whether a worksheet with the specified name already exists in the workbook
    private static bool WorksheetExists(Workbook wb, string name)
    {
        foreach (Worksheet ws in wb.Worksheets)
        {
            if (string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }
}
