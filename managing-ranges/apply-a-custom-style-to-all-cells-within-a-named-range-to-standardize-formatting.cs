// Title: Apply a bold yellow centered style to all cells in a named range with Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an existing Excel file, fetches a named range, creates a style with bold font, yellow fill, and centered alignment, and applies it to every cell using Aspose.Cells. | Write a reusable C# method that takes a Workbook, a named range name, and a Style object, then applies the provided style to the entire range with Aspose.Cells. | Modify the sample to use a blue background and italic font instead of bold yellow while keeping the same named range formatting.
// Common Searches: Aspose.Cells C# apply custom style to a defined name range | How to set bold font and yellow fill for all cells in a named range using Aspose.Cells .NET | Iterate through cells of a named range and change alignment with Aspose.Cells | Reusable function to style named ranges in Excel with Aspose.Cells for .NET | Change background color of a named range programmatically in C# Aspose.Cells
// Tags: apply custom style to named range Aspose.Cells | set bold font and yellow fill Aspose.Cells | center alignment for range cells .NET | create reusable style method Aspose.Cells | named range formatting Excel C#

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The program loads 'input.xlsx', retrieves the named range 'MyRange', creates a bold yellow centered style, applies it to each cell in the range, and saves the result as 'output.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range
            Name namedRange = workbook.Worksheets.Names["MyRange"];
            if (namedRange == null)
            {
                Console.WriteLine("Named range 'MyRange' not found.");
                return;
            }

            // Get the actual cell range that the name refers to
            Aspose.Cells.Range range = namedRange.GetRange();

            // Create a custom style
            Style customStyle = workbook.CreateStyle();
            customStyle.Font.IsBold = true;
            customStyle.ForegroundColor = Color.Yellow;
            customStyle.Pattern = BackgroundType.Solid;
            customStyle.HorizontalAlignment = TextAlignmentType.Center;
            customStyle.VerticalAlignment = TextAlignmentType.Center;

            // Apply the custom style to each cell in the range
            foreach (Cell cell in range)
            {
                cell.SetStyle(customStyle);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
