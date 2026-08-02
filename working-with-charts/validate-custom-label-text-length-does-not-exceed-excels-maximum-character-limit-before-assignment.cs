// Title: Validate and Truncate Label Shape Text (32,767 chars) with Aspose.Cells for .NET
// Description: Creates a workbook, adds a label shape, checks a string against Excel's 32,767‑character limit for shape text, truncates if necessary, optionally disables restriction checking, and saves the file.
// Keywords: Aspose.Cells label text limit | Excel shape text 32767 characters | truncate label text Aspose | validate shape text length .NET | CheckExcelRestriction Aspose.Cells
// Common Searches: Aspose.Cells label shape text length limit | how to truncate long text for Excel shape using Aspose | disable Excel restriction checking Aspose.Cells | maximum characters for a shape in Excel .NET
// Developer Intent: Ensure label shape text does not exceed Excel's 32,767‑character limit before assignment.
// Use Cases: Validate user‑generated strings before adding them to a label shape to avoid runtime errors. | Automatically shorten oversized text when generating dynamic reports with Aspose.Cells. | Turn off Excel restriction checks when storing longer text for later processing.
// AI Prompts: Write C# code with Aspose.Cells that checks a string's length against the 32,667‑character shape limit and truncates it if needed. | Show how to disable Excel restriction checking in Aspose.Cells while still saving the workbook. | Provide an example that validates and truncates text for multiple label shapes in a single worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsLabelLengthValidation
{
    // Creates a workbook, adds a label shape, checks a string against Excel's 32,767‑character limit for shape text, truncates if necessary, optionally disables restriction checking, and saves the file.
    class Program
    {
        // Excel's maximum characters for a shape's text (including labels) is 32,767.
        const int MaxLabelTextLength = 32767;

        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a label shape to the worksheet
            // Parameters: upper left row, upper left column, top, left, height, width (in pixels)
            Label labelShape = sheet.Shapes.AddLabel(2, 2, 50, 50, 200, 100);

            // Example of a long string that exceeds Excel's limit
            string longText = new string('A', 40000); // 40,000 characters

            // Validate length before assigning to the label
            if (longText.Length > MaxLabelTextLength)
            {
                // Option 1: Truncate the text to the allowed maximum
                string truncated = longText.Substring(0, MaxLabelTextLength);
                labelShape.Text = truncated;

                Console.WriteLine($"Input text was too long ({longText.Length} chars). Truncated to {MaxLabelTextLength} chars.");
            }
            else
            {
                // Length is within the limit; assign directly
                labelShape.Text = longText;
                Console.WriteLine($"Input text assigned successfully ({longText.Length} chars).");
            }

            // Optional: Disable Excel restriction checking if you need to store longer text
            // workbook.Settings.CheckExcelRestriction = false;

            // Save the workbook (lifecycle rule: save)
            workbook.Save("LabelWithValidatedText.xlsx");
        }
    }
}
