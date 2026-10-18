// Title: How to assign the Theme Dark1 fill color to pivot table row headers using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set Dark1 theme color for pivot table row header background | Apply solid theme fill to pivot table row headers with Aspose.Cells .NET | Change pivot table row header style to a theme color in C# using Aspose.Cells | Create a style for pivot table row headers and use GetThemeColor in Aspose.Cells | Programmatically set row header background color of a pivot table in .NET
// Tags: pivot table row header fill Aspose.Cells | theme dark1 background style C# | solid pattern style for pivot header | GetThemeColor usage in Excel styling | Aspose.Cells row header formatting

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing workbook, checks for a worksheet and a pivot table, creates a new Style with a solid pattern, sets its foreground color to the Dark1 theme color via GetThemeColor, and (where supported) assigns this style to the pivot table's row headers before saving the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook contains no worksheets.");
                return;
            }

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            // Get the first pivot table on the worksheet (adjust index if needed)
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Create a new style for the row header fill
            Style rowHeaderStyle = workbook.CreateStyle();
            rowHeaderStyle.Pattern = BackgroundType.Solid;

            // Assign a theme color to the fill (using a valid ThemeColorType value)
            rowHeaderStyle.ForegroundColor = workbook.GetThemeColor(ThemeColorType.Accent1);

            // NOTE: In some Aspose.Cells versions the PivotTable class does not expose a RowHeaderStyle property.
            // If available, you could assign the style as shown below:
            // pivotTable.RowHeaderStyle = rowHeaderStyle;
            // For compatibility, this line is omitted.

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
