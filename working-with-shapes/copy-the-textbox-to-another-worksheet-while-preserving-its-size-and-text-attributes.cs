// Title: Copy a textbox shape from one worksheet to another while preserving size and formatting with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy textbox shape to another worksheet preserving dimensions | How to duplicate a textbox shape between sheets using AddCopy in Aspose.Cells | Maintain original position and text formatting when moving shapes in Excel with Aspose.Cells | Copy shape from Sheet1 to Sheet2 preserving size in Aspose.Cells for .NET | C# Aspose.Cells copy textbox shape with same row/column bounds
// Tags: AddCopy method shape cloning Aspose.Cells | preserve shape dimensions during worksheet copy | retain textbox formatting after shape duplication | Aspose.Cells shape duplication across worksheets | maintain shape coordinates after copy

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing workbook, identifies a textbox shape on Sheet1, and uses the AddCopy method to duplicate it on Sheet2 while keeping the original row/column bounds. It then aligns the copied shape's Left and Top properties to match the source, ensuring the size, position, and text attributes are preserved before saving the workbook.
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
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get source and destination worksheets
            Worksheet srcSheet = workbook.Worksheets["Sheet1"];
            Worksheet destSheet = workbook.Worksheets["Sheet2"];

            // Locate the first shape in the source sheet (assumed to be a textbox)
            Shape srcTextbox = null;
            foreach (Shape shape in srcSheet.Shapes)
            {
                srcTextbox = shape;
                break;
            }

            if (srcTextbox != null)
            {
                // Copy the shape to the destination sheet preserving size and position
                Shape copiedShape = destSheet.Shapes.AddCopy(
                    srcTextbox,
                    srcTextbox.UpperLeftRow,
                    srcTextbox.UpperLeftColumn,
                    srcTextbox.LowerRightRow,
                    srcTextbox.LowerRightColumn);

                // Preserve the original position (optional)
                copiedShape.Left = srcTextbox.Left;
                copiedShape.Top = srcTextbox.Top;
            }
            else
            {
                Console.WriteLine("No shape found on Sheet1.");
            }

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
