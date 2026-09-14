// Title: Create a 45‑degree text rotation style and apply it to a vertical header range (A1:A10) using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a Style with RotationAngle = 45 and applies it to the range A1:A10 in an Aspose.Cells workbook. | Demonstrate how to set StyleFlag.All to true so that all style attributes, including rotation, are applied to a column range with Aspose.Cells. | Provide a complete example that saves the workbook after styling the vertical header with a 45‑degree rotated text style.
// Common Searches: Aspose.Cells C# set text rotation for a column header | apply custom style with rotation angle to range A1:A10 Aspose.Cells | rotate header text 45 degrees using Aspose.Cells .NET API | how to use StyleFlag to apply full style to a range in Aspose.Cells
// Tags: Aspose.Cells style rotation angle | apply style to range Aspose.Cells | StyleFlag all attributes Aspose.Cells | vertical header formatting Aspose.Cells | C# Aspose.Cells rotated header

using System;
using System.IO;
using Aspose.Cells;

// Shows how to create a Style with a 45‑degree RotationAngle, use StyleFlag.All to copy all style properties, apply the style to the vertical header range A1:A10, and save the workbook as RotatedHeader.xlsx with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a style object
            Style rotatedStyle = workbook.CreateStyle();

            // Set text rotation to 45 degrees (use RotationAngle property)
            rotatedStyle.RotationAngle = 45;

            // Apply all style attributes
            StyleFlag flag = new StyleFlag { All = true };

            // Define the vertical header range (e.g., A1:A10)
            Aspose.Cells.Range headerRange = sheet.Cells.CreateRange("A1:A10");

            // Apply the rotated style to the header range
            headerRange.ApplyStyle(rotatedStyle, flag);

            // Determine output file path
            string outputPath = "RotatedHeader.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
