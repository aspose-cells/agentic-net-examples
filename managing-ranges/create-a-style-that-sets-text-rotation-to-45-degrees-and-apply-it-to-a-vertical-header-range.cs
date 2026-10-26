// Title: Create a 45-degree rotated text style and apply it to a vertical header range with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set cell text rotation to 45 degrees | How to apply a custom style to a range of cells in Aspose.Cells .NET | Create a rotated header column in Excel using Aspose.Cells for C# | Apply StyleFlag.All when formatting a range with Aspose.Cells | Set vertical header formatting with rotation in Aspose.Cells C#
// Tags: Style.RotationAngle Aspose.Cells | StyleFlag.All property Aspose.Cells | ApplyStyle method range Aspose.Cells | vertical header range formatting Aspose.Cells | rotated text style Excel .NET | C# Aspose.Cells style rotation

using Aspose.Cells;
using System;

// The example creates a Workbook, defines a Style with a 45-degree RotationAngle, enables all style attributes via StyleFlag.All, creates a vertical header range (A1:A5), applies the rotated style to that range, populates the header cells, and saves the file as RotatedHeader.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Create a style and set text rotation to 45 degrees
            Style rotatedStyle = workbook.CreateStyle();
            rotatedStyle.RotationAngle = 45; // rotation in degrees

            // Define a style flag to apply the rotation attribute (apply all to be safe)
            StyleFlag flag = new StyleFlag();
            flag.All = true; // apply all style attributes, including rotation

            // Define the vertical header range (e.g., cells A1 to A5)
            Aspose.Cells.Range headerRange = sheet.Cells.CreateRange("A1:A5");

            // Apply the rotated style to the header range
            headerRange.ApplyStyle(rotatedStyle, flag);

            // Populate the header cells with sample text
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue($"Header {i + 1}");
            }

            // Save the workbook to a file
            workbook.Save("RotatedHeader.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
