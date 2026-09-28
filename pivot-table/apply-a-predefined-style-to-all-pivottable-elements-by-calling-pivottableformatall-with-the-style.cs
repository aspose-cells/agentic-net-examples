// Title: Apply a custom light‑blue style to every PivotTable in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an existing .xlsx file with Aspose.Cells, creates a solid light‑blue style, and applies it to all PivotTables using PivotTable.FormatAll. | Generate a method that iterates through each worksheet in a workbook, formats each PivotTable with a predefined style, and saves the updated file.
// Common Searches: Aspose.Cells C# how to apply a style to all pivot tables in a workbook | C# PivotTable.FormatAll example with custom style | set background color for every pivot table using Aspose.Cells | iterate worksheets and format pivot tables programmatically in .NET
// Tags: PivotTable.FormatAll custom style Aspose.Cells | apply solid background to pivot tables .NET | create workbook style programmatically Aspose.Cells | format all pivot tables in Excel using C# | load and save workbook with styled pivot tables Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// // Loads Input.xlsx, creates a light‑blue solid background style, loops through each worksheet and its PivotTables, applies the style to the entire pivot table via PivotTable.FormatAll, and saves the result as Output.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "Input.xlsx";
        const string outputPath = "Output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Create a custom style (as a fallback when built‑in pivot styles are unavailable)
            Style pivotStyle = workbook.CreateStyle();
            pivotStyle.ForegroundColor = Color.LightBlue;
            pivotStyle.Pattern = BackgroundType.Solid;

            // Iterate through all worksheets and their pivot tables
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    // Apply the style to the entire pivot table
                    pivotTable.FormatAll(pivotStyle);
                }
            }

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
