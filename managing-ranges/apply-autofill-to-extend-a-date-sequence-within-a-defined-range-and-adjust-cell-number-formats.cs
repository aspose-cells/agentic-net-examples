// Title: C# – AutoFill Date Series and Apply Built‑in Date Format to a Column with Aspose.Cells
// Description: Creates a workbook, seeds two dates, uses AutoFill (Series) to extend the sequence, applies the built‑in date format (Number 14) to the whole column via StyleFlag, auto‑fits the column, and saves the file.
// Keywords: Aspose.Cells AutoFill | date series C# | built‑in date format | StyleFlag number format | auto fit column Aspose.Cells | C# workbook date formatting | Aspose.Cells range AutoFill Series
// Common Searches: Aspose.Cells AutoFill date series example | How to set built‑in date format for a column in Aspose.Cells | Using StyleFlag to apply number format only in Aspose.Cells | Auto‑fit column width after formatting cells Aspose.Cells | C# generate sequential dates with Aspose.Cells
// Developer Intent: Create a column of sequential dates and format it with a built‑in date style using Aspose.Cells for .NET.
// Use Cases: Generate a daily schedule worksheet by auto‑filling dates and displaying them in mm‑dd‑yy format. | Prepare a financial report where the date column must stay consistent across many rows while automatically adjusting column width. | Export a project timeline with auto‑filled dates and a uniform column style for easy readability.
// AI Prompts: Write C# code with Aspose.Cells to auto‑fill a date series from A1:A2 to A3:A20 and set the column to built‑in date format 14. | Explain how StyleFlag determines which style attributes are applied when formatting cells in Aspose.Cells. | Provide a step‑by‑step guide to auto‑fill a range with a series, apply a number format to an entire column, and auto‑fit the column width.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, seeds two dates, uses AutoFill (Series) to extend the sequence, applies the built‑in date format (Number 14) to the whole column via StyleFlag, auto‑fits the column, and saves the file.
public class AutoFillDateSeriesDemo
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Seed the first two cells with dates to define the series
        cells["A1"].PutValue(new DateTime(2023, 1, 1));
        cells["A2"].PutValue(new DateTime(2023, 1, 2));

        // Define the source and target ranges (use fully qualified Aspose.Cells.Range)
        Aspose.Cells.Range sourceRange = cells.CreateRange("A1:A2");
        Aspose.Cells.Range targetRange = cells.CreateRange("A3:A10");

        // Extend the date sequence using AutoFill with the Series type
        sourceRange.AutoFill(targetRange, AutoFillType.Series);

        // Create a style that uses a built‑in date number format (e.g., mm-dd-yy)
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Number = 14; // Built‑in date format

        // Prepare a StyleFlag to indicate which style attributes to apply (Number format only)
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true; // Apply number format

        // Apply the date style to the entire column A (column index 0)
        cells.ApplyColumnStyle(0, dateStyle, flag);

        // Auto‑fit column A to display the dates properly
        sheet.AutoFitColumns(0, 0);

        // Determine output path and ensure the directory exists
        string outputPath = Path.Combine(Environment.CurrentDirectory, "DateSeriesAutoFill.xlsx");
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception saveEx)
        {
            Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
        }
    }
}
