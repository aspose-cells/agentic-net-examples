// Title: Validate and Truncate Label Text Length for Excel Shapes with Aspose.Cells (.NET)
// Description: Shows how to create a workbook, add a label shape, verify the text against Excel's 32,767‑character limit, truncate or raise an error if exceeded, assign the safe text, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | label shape | text length limit | truncate label text | Excel maximum characters | shape text validation | workbook | Excel shape label
// Common Searches: Aspose.Cells limit label text length | truncate Excel shape label Aspose | maximum characters for Excel shape label | validate label text before saving Aspose.Cells | C# check label text length Aspose.Cells
// Developer Intent: Check and enforce Excel's 32,767‑character limit on shape label text before assigning it in a .NET application.
// Use Cases: Prevent runtime exceptions when generating reports with overly long annotations. | Automatically shorten user‑provided comments in dashboards or chart labels. | Provide a reusable validation method for any shape (label, textbox) in Aspose.Cells. | Enforce corporate text‑length policies during data export processes.
// AI Prompts: Generate C# code using Aspose.Cells that adds a label shape and trims its text to 32767 characters. | Show how to throw an ArgumentException when label text exceeds Excel's limit in Aspose.Cells. | Create a helper method ValidateLabelText(string text) that returns a safe string for any shape. | Write a unit test that verifies label text truncation logic with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsLabelLengthValidation
{
    // Shows how to create a workbook, add a label shape, verify the text against Excel's 32,767‑character limit, truncate or raise an error if exceeded, assign the safe text, and save the file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a label shape to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, height, width
            int upperLeftRow = 2;
            int upperLeftColumn = 2;
            int top = 10;
            int left = 10;
            int height = 100;
            int width = 300;
            Label label = sheet.Shapes.AddLabel(upperLeftRow, upperLeftColumn, top, left, height, width);

            // Text to assign to the label
            string labelText = new string('A', 35000); // Example long text

            // Excel's maximum character limit for a cell/label text is 32,767 characters
            const int ExcelMaxTextLength = 32767;

            // Validate length before assignment
            if (labelText.Length > ExcelMaxTextLength)
            {
                // Option 1: Truncate the text to the maximum allowed length
                labelText = labelText.Substring(0, ExcelMaxTextLength);
                // Optionally, you could throw an exception instead:
                // throw new ArgumentException($"Label text exceeds Excel's maximum length of {ExcelMaxTextLength} characters.");
            }

            // Assign the validated (or truncated) text to the label
            label.Text = labelText;

            // Save the workbook (save rule)
            string outputPath = "LabelWithValidatedText.xlsx";
            workbook.Save(outputPath);
        }
    }
}
