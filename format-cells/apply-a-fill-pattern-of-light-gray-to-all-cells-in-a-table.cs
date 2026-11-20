// Title: How to apply a light gray solid fill to a specific table range (A1:D10) in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: apply solid light gray background to a range of cells with Aspose.Cells in C# | Aspose.Cells C# set background color for table range A1:D10 | use StyleFlag.All to format every cell in an Excel range with Aspose.Cells | create and apply a custom style to an Excel table using Aspose.Cells .NET | programmatically change the fill pattern of Excel cells to gray with Aspose.Cells
// Tags: apply solid fill to Excel range Aspose.Cells | set light gray background Aspose.Cells C# | StyleFlag All property Aspose.Cells | custom style for table cells Aspose.Cells | format table range background Aspose.Cells .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, defines the A1:D10 range as a table, creates a style with a LightGray foreground and solid fill pattern, applies the style to every cell in the range using a StyleFlag with All=true, and saves the modified file, handling missing input files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range that represents the table (e.g., A1:D10)
            var tableRange = sheet.Cells.CreateRange("A1:D10");

            // Create a style with a light gray fill pattern
            Style grayStyle = workbook.CreateStyle();
            grayStyle.ForegroundColor = Color.LightGray;
            grayStyle.Pattern = BackgroundType.Solid; // solid fill

            // Apply all style attributes
            StyleFlag flag = new StyleFlag { All = true };

            // Apply the style to every cell in the table range
            tableRange.ApplyStyle(grayStyle, flag);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
