// Title: Set worksheet FitToPagesWide dynamically from column count with a configurable factor using Aspose.Cells for .NET
// Common Searches: how to calculate FitToPagesWide from MaxDataColumn in Aspose.Cells C# | set page width based on number of columns and custom factor Aspose.Cells | adjust Excel worksheet pagination dynamically using Aspose.Cells .NET | use configurable factor to determine pages wide in Aspose.Cells page setup | auto fit to pages wide by column count Aspose.Cells C# example
// Tags: FitToPagesWide calculation using MaxDataColumn | configurable page width factor in Aspose.Cells | dynamic pagination for Excel worksheets .NET | Aspose.Cells page setup based on column count | C# Excel export page width adjustment

using Aspose.Cells;
using System;

// The example loads an Excel workbook, counts the data columns with MaxDataColumn, computes the required pages wide by dividing the column count by a configurable factor (rounded up, minimum 1), assigns this value to PageSetup.FitToPagesWide, sets FitToPagesTall to 0 for automatic height, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Configurable factor for page width calculation
        int factor = 5; // Change this value as needed

        // Load the workbook (replace with your actual file path)
        var workbook = new Workbook("input.xlsx");

        // Access the worksheet you want to modify (e.g., the first sheet)
        var sheet = workbook.Worksheets[0];

        // Determine the number of columns that contain data
        // MaxDataColumn is zero‑based, so add 1 to get the count
        int columnCount = sheet.Cells.MaxDataColumn + 1;

        // Calculate the number of pages wide based on the factor
        int pagesWide = (int)Math.Ceiling((double)columnCount / factor);
        if (pagesWide < 1) pagesWide = 1; // Ensure at least one page

        // Apply the FitToPagesWide setting
        var pageSetup = sheet.PageSetup;
        pageSetup.FitToPagesWide = pagesWide;
        // Allow the height to adjust automatically
        pageSetup.FitToPagesTall = 0;

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
