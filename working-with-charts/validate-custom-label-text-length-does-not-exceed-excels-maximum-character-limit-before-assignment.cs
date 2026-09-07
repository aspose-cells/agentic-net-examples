// Title: Validate and truncate a label shape's text to Excel's 32,767‑character limit using Aspose.Cells for .NET
// AI Prompts: Provide C# code that checks a label shape's Text length against Excel's 32,767‑character limit and truncates it before assignment with Aspose.Cells. | Show how to enable Excel restriction checking, add a label shape, validate its text size, and save the workbook using Aspose.Cells for .NET. | Demonstrate handling of overly long label text by truncating it and catching exceptions when saving a workbook with Aspose.Cells.
// Common Searches: Aspose.Cells how to enforce Excel string length limit on label shapes in C# | C# truncate label text exceeding 32767 characters before saving workbook with Aspose.Cells | validate shape text size in Aspose.Cells .NET to avoid Excel restriction errors | example of checking and trimming label text length using Aspose.Cells for .NET | Excel maximum characters for shape label Aspose.Cells validation code
// Tags: label shape text length validation Aspose.Cells | truncate label text Excel limit C# | check Excel restriction setting Aspose.Cells | shape text size enforcement Aspose.Cells | Excel string length limit handling Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, enables Excel restriction checking, adds a label shape, and validates the label's text against Excel's 32,667‑character limit. If the text exceeds the limit it is truncated, then the safe text is assigned to the label and the workbook is saved, with any errors captured and reported.
class ValidateLabelTextDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable Excel restriction checking (throws if limit exceeded)
            workbook.Settings.CheckExcelRestriction = true;

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a label shape to the worksheet
            // Parameters: upper left row, upper left column, row offset, column offset, height (pixels), width (pixels)
            // AddLabel returns a Label object in recent Aspose.Cells versions
            Label label = worksheet.Shapes.AddLabel(1, 1, 0, 0, 50, 200);

            if (label == null)
            {
                Console.WriteLine("Failed to create label shape.");
                return;
            }

            // Example text that may exceed Excel's maximum string length (32,767 characters)
            string labelText = new string('X', 35000); // 35,000 characters

            // Excel's maximum allowed length for a string in a cell/shape
            const int MaxExcelStringLength = 32767;

            // Validate length before assigning to the label
            if (labelText.Length > MaxExcelStringLength)
            {
                Console.WriteLine($"Label text length ({labelText.Length}) exceeds Excel limit. Truncating to {MaxExcelStringLength} characters.");
                labelText = labelText.Substring(0, MaxExcelStringLength);
            }

            // Assign the validated (or truncated) text to the label
            label.Text = labelText;

            // Save the workbook
            workbook.Save("LabelValidated.xlsx");
            Console.WriteLine("Workbook saved as 'LabelValidated.xlsx'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
