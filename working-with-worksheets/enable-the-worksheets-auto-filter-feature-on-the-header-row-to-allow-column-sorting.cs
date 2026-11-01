// Title: Enable AutoFilter on the header row of an Excel worksheet and save the file with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that adds an AutoFilter to the first row of a worksheet and saves the workbook using Aspose.Cells. | Create a reusable method that detects the used columns in any worksheet and applies an AutoFilter to that header row. | Generate a snippet that loads an existing Excel file, enables column sorting via AutoFilter on the header row, and overwrites the file.
// Common Searches: asp.net aspose.cells enable autofilter on first row c# example | c# set autofilter range dynamically based on used columns aspose.cells | how to apply column sorting filter in excel using aspose.cells c# | save workbook after adding autofilter with aspose.cells c#
// Tags: Aspose.Cells set AutoFilter range | C# enable worksheet autofilter | Aspose.Cells column sorting with AutoFilter | save Excel workbook after applying AutoFilter | dynamic AutoFilter range based on used columns

using Aspose.Cells;

// The program creates (or loads) a workbook, accesses the first worksheet, sets the AutoFilter range to the header row (e.g., A1:D1), and saves the workbook as an Excel file with the filter enabled.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // create

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Enable AutoFilter on the header row.
        // Adjust the range (e.g., "A1:D1") to match the actual header columns.
        sheet.AutoFilter.Range = "A1:D1";

        // Save the workbook with the AutoFilter applied
        workbook.Save("AutoFilterDemo.xlsx"); // save
    }
}
