// Title: How to apply a custom blue‑bold style to every PivotTable in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# apply same style to all pivot tables in a workbook | Set font color and background for every PivotTable element with PivotTable.FormatAll | C# code to format all pivot tables using a custom style in Aspose.Cells | Iterate worksheets and apply style to pivot tables Aspose.Cells .NET | Apply custom style to PivotTable elements programmatically with Aspose.Cells
// Tags: PivotTable.FormatAll custom style Aspose.Cells | apply consistent formatting to all pivot tables .NET | create workbook style with blue bold font Aspose.Cells | iterate worksheets and pivot table collections C# | programmatic pivot table element formatting Excel

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file, creates a style with a blue bold font and light‑yellow background, loops through each worksheet and its PivotTable collection, applies the style to all PivotTable elements via PivotTable.FormatAll, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Create a custom style
                Style style = workbook.CreateStyle();
                style.Font.Color = Color.Blue;
                style.Font.IsBold = true;
                style.Pattern = BackgroundType.Solid;
                style.ForegroundColor = Color.LightYellow;

                // Apply the style to every PivotTable in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Access the collection of pivot tables on the worksheet
                    PivotTableCollection pivotTables = sheet.PivotTables;

                    foreach (PivotTable pivotTable in pivotTables)
                    {
                        // Apply the style to all elements of the pivot table
                        pivotTable.FormatAll(style);
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
