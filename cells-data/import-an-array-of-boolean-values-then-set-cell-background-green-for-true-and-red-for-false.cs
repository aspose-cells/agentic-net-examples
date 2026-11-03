// Title: Write a bool[] to an Excel worksheet and color cells green for true and red for false using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# color cells green when boolean is true and red when false | How to apply conditional background colors to a boolean array in an Excel file with Aspose.Cells | Set cell background based on bool values using Aspose.Cells .NET example | Write boolean array to column and format cells with green/red colors in Aspose.Cells | C# Aspose.Cells example for coloring true/false cells in Excel
// Tags: write boolean array to worksheet Aspose.Cells | conditional cell background color Aspose.Cells | apply solid fill based on boolean Aspose.Cells | format true cells green Aspose.Cells | format false cells red Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The sample creates a workbook, writes each element of a bool[] into column A, and applies a solid green background for true values and a solid red background for false values before saving the file as BooleanColors.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Sample boolean array
            bool[] boolValues = new bool[] { true, false, true, true, false };

            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate column A with boolean values
            for (int i = 0; i < boolValues.Length; i++)
            {
                Cell cell = sheet.Cells[i, 0];
                cell.PutValue(boolValues[i]); // set the boolean value

                // Apply background color based on the boolean value
                Style style = cell.GetStyle();
                style.ForegroundColor = boolValues[i] ? Color.Green : Color.Red;
                style.Pattern = BackgroundType.Solid;
                cell.SetStyle(style);
            }

            // Ensure the output directory exists
            string outputPath = "BooleanColors.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
