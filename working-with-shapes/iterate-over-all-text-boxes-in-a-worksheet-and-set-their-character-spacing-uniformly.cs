// Title: C# – Set Uniform Character Spacing for All TextBoxes in an Aspose.Cells Worksheet
// Description: Demonstrates how to create a workbook, add TextBox shapes, define a spacing value, loop through the worksheet's TextBoxes collection, apply the same TextOptions.Spacing to each box, and save the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells TextBox spacing C# | uniform character spacing Aspose.Cells | TextOptions.Spacing property | iterate worksheet TextBoxes | .NET Excel shape formatting
// Common Searches: Aspose.Cells set character spacing for all text boxes C# | loop through TextBox collection in worksheet Aspose.Cells | apply uniform TextOptions.Spacing to Excel text boxes | C# code example Aspose.Cells TextBox formatting
// Developer Intent: Apply a single character‑spacing value to every TextBox in a worksheet programmatically.
// Use Cases: Ensure consistent typography across multiple text boxes in a report template. | Batch‑update existing Excel files to meet branding guidelines for text spacing. | Prepare a workbook for publishing where all text boxes need the same readability settings.
// AI Prompts: Write C# code using Aspose.Cells that iterates over all TextBox objects in a worksheet and sets TextOptions.Spacing to a given value. | Show how to read the current spacing of a TextBox, compare it with a target, and update only when different. | Explain best practices for bulk formatting of shape text properties in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsTextBoxSpacingDemo
{
    // Demonstrates how to create a workbook, add TextBox shapes, define a spacing value, loop through the worksheet's TextBoxes collection, apply the same TextOptions.Spacing to each box, and save the file with Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample text boxes for demonstration
            int tb1 = worksheet.TextBoxes.Add(1, 1, 150, 50);
            worksheet.TextBoxes[tb1].Text = "First TextBox";

            int tb2 = worksheet.TextBoxes.Add(3, 2, 200, 60);
            worksheet.TextBoxes[tb2].Text = "Second TextBox";

            // Define the uniform character spacing value
            double uniformSpacing = 2.0; // Adjust as needed

            // Iterate over all text boxes in the worksheet
            foreach (TextBox textBox in worksheet.TextBoxes)
            {
                // Set the character spacing uniformly using TextOptions.Spacing
                textBox.TextOptions.Spacing = uniformSpacing;
            }

            // Save the workbook (lifecycle save rule)
            workbook.Save("TextBoxUniformSpacing.xlsx");
        }
    }
}
