// Title: Aspose.Cells .NET: Create a 45° Text Rotation Style and Apply It to a Vertical Header Range
// Description: This example shows how to create a Workbook, define a Style with RotationAngle = 45°, enable rotation via StyleFlag, apply the style to the range A1:A5 (vertical header), populate sample text, and save the file as HeaderRotationExample.xlsx.
// Keywords: Aspose.Cells text rotation | C# rotate text 45 degrees | StyleFlag rotation | apply style to range | vertical header style | Excel header angle | Aspose.Cells .NET example
// Common Searches: Aspose.Cells rotate text 45 degrees | How to apply rotated style to a range in C# | Set text rotation for column header using Aspose.Cells | Enable text rotation with StyleFlag in Aspose.Cells | Create angled header cells in Excel with Aspose
// Developer Intent: Create a reusable style that rotates cell text 45° and apply it to a vertical header range in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Design compact tables where column headings are angled to save width. | Generate reports with visually distinct, slanted header labels. | Automate Excel exports that require rotated text for better readability.
// AI Prompts: Generate code to change the rotation angle based on a user‑provided value in Aspose.Cells. | Show how to apply the same 45° rotation style to multiple non‑contiguous ranges in a workbook. | Explain how to combine text rotation with font styling and background color in a single Aspose.Cells style.

using System;
using Aspose.Cells;

namespace AsposeCellsStyleRotationExample
{
    // This example shows how to create a Workbook, define a Style with RotationAngle = 45°, enable rotation via StyleFlag, apply the style to the range A1:A5 (vertical header), populate sample text, and save the file as HeaderRotationExample.xlsx.
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

                // Enable the rotation setting using a style flag
                StyleFlag styleFlag = new StyleFlag();
                styleFlag.Rotation = true;

                // Define the vertical header range (e.g., cells A1 to A5)
                Aspose.Cells.Range headerRange = worksheet.Cells.CreateRange("A1:A5");

                // Apply the style with the rotation flag to the range
                headerRange.ApplyStyle(rotationStyle, styleFlag);

                // Optionally put some sample text in the header cells to see the effect
                for (int row = 0; row < 5; row++)
                {
                    worksheet.Cells[row, 0].PutValue($"Header {row + 1}");
                }

                // Save the workbook
                workbook.Save("HeaderRotationExample.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
