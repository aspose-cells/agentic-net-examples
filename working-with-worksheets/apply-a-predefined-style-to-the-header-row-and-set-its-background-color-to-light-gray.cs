// Title: Apply a predefined style with a light gray background to the header row of an Excel worksheet using Aspose.Cells for .NET
// Common Searches: how to apply a solid light gray fill to the first row in an Excel file with Aspose.Cells C# | setting bold font and background color for header row using Aspose.Cells .NET API | create and apply a style to the header range in Aspose.Cells workbook | using StyleFlag to format an entire header row in C# with Aspose.Cells | programmatically style the top row of an Excel worksheet in .NET
// Tags: Aspose.Cells header row style creation | C# background color for Excel row Aspose.Cells | StyleFlag solid fill usage Aspose.Cells | custom style with light gray fill Aspose.Cells | formatting first worksheet row .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads or creates a workbook, defines a style with a solid light gray background and bold font, determines the used column count, creates a range covering the first row, applies the style to that range using a StyleFlag with All set to true, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            Worksheet sheet = workbook.Worksheets[0];

            // Create a style for the header row
            Style headerStyle = workbook.CreateStyle();
            headerStyle.ForegroundColor = Color.LightGray;   // background color
            headerStyle.Pattern = BackgroundType.Solid;     // apply solid fill
            headerStyle.Font.IsBold = true;                 // make header text bold

            // Determine the range of the header row (first row, all used columns)
            int firstRow = 0;
            int firstColumn = 0;
            int totalColumns = sheet.Cells.MaxColumn + 1;   // include all columns with data

            // Apply the style to the header row
            Aspose.Cells.Range headerRange = sheet.Cells.CreateRange(firstRow, firstColumn, 1, totalColumns);
            StyleFlag flag = new StyleFlag { All = true };
            headerRange.ApplyStyle(headerStyle, flag);

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
