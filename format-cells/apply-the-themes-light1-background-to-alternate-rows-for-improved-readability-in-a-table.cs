// Title: How to apply a Light Gray background to every other row in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example to shade alternate rows with a light gray fill | apply alternating row background color in Excel using Aspose.Cells .NET API | how to use StyleFlag and Range to format every second row in a workbook with Aspose.Cells | set solid light gray background for odd rows in an Excel file with Aspose.Cells C# | Aspose.Cells .NET code to improve table readability by coloring alternate rows
// Tags: alternating row background Aspose.Cells .NET | solid light gray style Aspose.Cells | apply style to range Aspose.Cells C# | StyleFlag usage for row formatting | Excel table readability formatting Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The sample loads (or creates) an Excel workbook, defines a solid light‑gray style, and iterates through the used rows, applying the style to every second row via a Range and a StyleFlag, then saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                var tempSheet = tempWb.Worksheets[0];
                tempSheet.Cells["A1"].PutValue("Sample Data");
                tempWb.Save(inputPath);
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Create a style with a solid light gray background.
            Style lightStyle = workbook.CreateStyle();
            lightStyle.ForegroundColor = Color.LightGray; // concrete color
            lightStyle.Pattern = BackgroundType.Solid;    // solid fill

            // Apply all style attributes.
            StyleFlag flag = new StyleFlag { All = true };

            // Determine the used range of the sheet.
            int maxRow = sheet.Cells.MaxDataRow;          // zero‑based index of last used row
            int maxColumn = sheet.Cells.MaxDataColumn;    // zero‑based index of last used column

            // Apply the background to every alternate row (e.g., rows 0,2,4,…).
            for (int row = 0; row <= maxRow; row += 2)
            {
                // Create a range that spans the entire row within the used columns.
                AsposeRange rowRange = sheet.Cells.CreateRange(row, 0, 1, maxColumn + 1);
                // Apply the style to the range.
                rowRange.ApplyStyle(lightStyle, flag);
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
