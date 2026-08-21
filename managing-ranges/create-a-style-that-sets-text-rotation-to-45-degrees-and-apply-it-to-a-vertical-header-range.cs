// Title: Aspose.Cells .NET: Create a 45° Text Rotation Style and Apply to a Vertical Header Range (A1:A5)
// Description: Demonstrates how to create a reusable Style with a 45‑degree RotationAngle, enable the rotation flag, define the A1:A5 header range, apply the style to that range, and save the workbook as VerticalHeaderWithRotation.xlsx using C# and Aspose.Cells.
// Keywords: Aspose.Cells | .NET | C# | text rotation | RotationAngle | StyleFlag | apply style to range | vertical header | Excel export | A1:A5
// Common Searches: Aspose.Cells rotate text 45 degrees | apply rotation style to a range Aspose.Cells .NET | set text orientation for column header in Aspose.Cells | how to use StyleFlag for rotation in C# | create rotated header cells with Aspose.Cells
// Developer Intent: Create a 45° text‑rotation style and apply it to the vertical header range A1:A5 in an Aspose.Cells workbook.
// Use Cases: Design reports where column headers are tilted to save horizontal space. | Export spreadsheets with vertically oriented titles for narrow columns while preserving other formatting. | Reuse a single rotation style across multiple header ranges in large workbooks.
// AI Prompts: Generate a method that accepts any cell range and applies a 45° rotation style using Aspose.Cells. | Show code to let users specify the rotation angle at runtime while keeping existing style attributes intact. | Explain how to combine text rotation with font, border, and background styles when applying to a range in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsStyleRotationExample
{
    // Demonstrates how to create a reusable Style with a 45‑degree RotationAngle, enable the rotation flag, define the A1:A5 header range, apply the style to that range, and save the workbook as VerticalHeaderWithRotation.xlsx using C# and Aspose.Cells.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Create a style and set the text rotation angle to 45 degrees
                Style rotationStyle = workbook.CreateStyle();
                rotationStyle.RotationAngle = 45;

                // Enable the rotation setting via a style flag
                StyleFlag styleFlag = new StyleFlag();
                styleFlag.Rotation = true;

                // Define the vertical header range (first column, rows 1 to 5)
                // Use fully qualified type to avoid conflict with System.Range
                Aspose.Cells.Range headerRange = worksheet.Cells.CreateRange("A1:A5");

                // Apply the style with the rotation flag to the range
                headerRange.ApplyStyle(rotationStyle, styleFlag);

                // Save the workbook
                workbook.Save("VerticalHeaderWithRotation.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
