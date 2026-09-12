// Title: Apply a light gray solid fill pattern to all cells in a specific table range (A1:D10) using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a Style with a solid light gray fill and applies it to the range A1:D10 in an Aspose.Cells workbook. | Write a method that loads an Excel file (or creates a new workbook if missing), defines a table range, and sets a light gray background for every cell using Aspose.Cells for .NET. | Provide a reusable function that accepts an Aspose.Cells Range object and applies a solid light gray background style to all its cells.
// Common Searches: how to set a solid light gray background for a range of cells in Aspose.Cells C# | Aspose.Cells apply fill color to entire table A1:D10 | C# Aspose.Cells style cells with light gray pattern | set background color for Excel table using Aspose.Cells .NET API
// Tags: Aspose.Cells apply solid fill to range | C# set cell background color Aspose.Cells | light gray fill pattern Excel Aspose | style table range Aspose.Cells .NET | create style and apply to cells Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// // Loads an existing workbook or creates a new one, defines the table range A1:D10, creates a style with a solid light gray fill, applies the style to all cells in the range, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; if not, create a new workbook with a default sheet.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                // Optionally add some sample data to avoid an empty sheet.
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue("Sample");
            }

            // Access the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range representing the table (A1:D10).
            AsposeRange tableRange = worksheet.Cells.CreateRange("A1", "D10");

            // Create a new style for the fill pattern.
            Style fillStyle = workbook.CreateStyle();
            fillStyle.Pattern = BackgroundType.Solid;          // Solid fill.
            fillStyle.ForegroundColor = Color.LightGray;      // Light gray color.
            fillStyle.BackgroundColor = Color.White;          // Optional background color.

            // Apply the style to all cells in the range.
            StyleFlag styleFlag = new StyleFlag { All = true };
            tableRange.ApplyStyle(fillStyle, styleFlag);

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
