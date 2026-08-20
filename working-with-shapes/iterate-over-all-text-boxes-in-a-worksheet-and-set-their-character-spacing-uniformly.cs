// Title: Set Uniform Character Spacing for All TextBox Shapes in Aspose.Cells (C#)
// Description: Demonstrates how to create or load a workbook, add TextBox shapes, define a spacing value, loop through the worksheet's TextBoxes collection, apply the same TextOptions.Spacing to each box, and save the Excel file.
// Keywords: Aspose.Cells C# TextBox spacing | character spacing Aspose.Cells | TextOptions.Spacing .NET | iterate TextBox collection | uniform shape formatting Excel | Aspose.Cells shape properties
// Common Searches: how to change character spacing for all text boxes in Aspose.Cells C# | set same spacing for multiple TextBox shapes in Excel using Aspose.Cells | Aspose.Cells iterate TextBoxes and modify TextOptions | C# code to apply uniform text spacing in a workbook | Aspose.Cells TextBox formatting example
// Developer Intent: Apply a single character‑spacing value to every TextBox shape in a worksheet.
// Use Cases: Ensure consistent appearance of callout boxes in automatically generated reports. | Enforce brand‑compliant text layout across dashboard widgets. | Prepare a template where all comment boxes share identical spacing before data insertion.
// AI Prompts: Write C# code that loads an existing workbook and sets TextOptions.Spacing to 1.5 for all TextBox objects. | Show how to increase spacing only for TextBoxes containing the word "Important" while leaving others unchanged. | Explain how to read the current spacing of a TextBox, calculate a 20% increase, and apply it to every TextBox in the sheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsTextBoxSpacingDemo
{
    // Demonstrates how to create or load a workbook, add TextBox shapes, define a spacing value, loop through the worksheet's TextBoxes collection, apply the same TextOptions.Spacing to each box, and save the Excel file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a few sample text boxes for demonstration
            int tbIndex1 = worksheet.TextBoxes.Add(1, 1, 150, 50);
            TextBox textBox1 = worksheet.TextBoxes[tbIndex1];
            textBox1.Text = "First TextBox";

            int tbIndex2 = worksheet.TextBoxes.Add(3, 2, 200, 60);
            TextBox textBox2 = worksheet.TextBoxes[tbIndex2];
            textBox2.Text = "Second TextBox";

            // Define the uniform character spacing value
            double uniformSpacing = 2.0; // Positive value increases spacing, negative decreases

            // Iterate over all TextBox objects in the worksheet and set the spacing
            foreach (TextBox tb in worksheet.TextBoxes)
            {
                // TextOptions provides access to character spacing
                tb.TextOptions.Spacing = uniformSpacing;
            }

            // Save the workbook to a file
            workbook.Save("TextBoxUniformSpacing.xlsx");
        }
    }
}
