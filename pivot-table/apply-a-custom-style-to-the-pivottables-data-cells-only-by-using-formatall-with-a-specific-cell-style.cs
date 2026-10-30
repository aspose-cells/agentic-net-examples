// Title: C# example: Apply a custom style to all data cells of an Excel PivotTable using Aspose.Cells FormatAll
// Common Searches: Aspose.Cells C# how to use FormatAll to style pivot table data cells | apply custom background and font to pivot table values with Aspose.Cells .NET | C# code sample for formatting all pivot table data cells in an Excel workbook | set light yellow fill and bold dark blue font for pivot table data using Aspose.Cells | formatting pivot table data area programmatically in .NET with Aspose.Cells
// Tags: pivot table custom style aspnet | excel pivot table cell formatting aspnet | c# apply style to pivot data cells | aspnet workbook pivot table styling | c# excel pivot table formatting example

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Ensure PivotTable types are available

// The program loads an existing Excel workbook, accesses its first worksheet and PivotTable, creates a custom style with a light‑yellow fill and bold dark‑blue font, applies this style to every data cell of the PivotTable using the FormatAll method, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the first PivotTable in the worksheet
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No PivotTables found in the worksheet.");
                return;
            }
            PivotTable pivotTable = sheet.PivotTables[0];

            // Create a custom style for the PivotTable data cells
            Style customStyle = workbook.CreateStyle();
            customStyle.ForegroundColor = Color.LightYellow;
            customStyle.Pattern = BackgroundType.Solid;
            customStyle.Font.IsBold = true;
            customStyle.Font.Color = Color.DarkBlue;

            // Apply the custom style to all data cells of the PivotTable
            pivotTable.FormatAll(customStyle);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
