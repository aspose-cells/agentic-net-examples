// Title: C# Example: Validate and Clamp TextBox Character Spacing (-10 to 10) with Aspose.Cells
// Description: This Aspose.Cells for .NET sample creates a workbook, adds a TextBox shape, sets its TextOptions.Spacing property, checks whether the value lies between -10 and 10 points, automatically clamps out‑of‑range values to the nearest limit, and saves the file. It demonstrates proper handling of character‑spacing constraints for shape text.
// Keywords: Aspose.Cells TextBox spacing | C# TextOptions.Spacing range | character spacing validation .NET | clamp shape spacing Aspose | Aspose.Cells example GitHub | Excel shape text spacing | adjust TextBox character spacing | Aspose.Cells API usage | range check -10 to 10 points | C# workbook shape validation
// Common Searches: How to limit TextBox character spacing in Aspose.Cells? | Validate TextOptions.Spacing range in C# | Clamp out‑of‑range spacing for Excel shapes using Aspose | What is the allowed spacing range for Aspose.Cells TextBox? | Example code for correcting shape spacing in .NET
// Developer Intent: Ensure a TextBox's character spacing stays within the supported -10 to 10 point interval and automatically correct values that fall outside this range.
// Use Cases: Sanitize user‑provided spacing values before applying them to report templates. | Batch‑process worksheets to enforce spacing limits on all TextBox shapes. | Integrate spacing validation into a CI pipeline that generates Excel files with dynamic text styling.
// AI Prompts: Generate a C# method that receives a TextOptions object and clamps its Spacing property to the -10..10 point range. | Write code that iterates over every shape in a worksheet and ensures each TextBox's character spacing complies with Aspose.Cells limits. | Provide a logging snippet that warns when a TextBox spacing value is out of range, then corrects it before saving the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// This Aspose.Cells for .NET sample creates a workbook, adds a TextBox shape, sets its TextOptions.Spacing property, checks whether the value lies between -10 and 10 points, automatically clamps out‑of‑range values to the nearest limit, and saves the file. It demonstrates proper handling of character‑spacing constraints for shape text.
class ValidateCharacterSpacing
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        TextBox textBox = sheet.Shapes.AddTextBox(0, 0, 2, 0, 200, 100);
        textBox.Text = "Sample Text";

        // Access the TextOptions of the text box
        TextOptions textOptions = textBox.TextOptions;

        // Set a spacing value (example value that may be out of range)
        textOptions.Spacing = 12.5; // Points

        // Validate that the spacing is within the allowed range of -10 to 10 points
        double spacing = textOptions.Spacing;
        if (spacing < -10.0 || spacing > 10.0)
        {
            // Adjust to the nearest allowed value
            double corrected = Math.Max(-10.0, Math.Min(10.0, spacing));
            Console.WriteLine($"Spacing {spacing} is out of range. Adjusting to {corrected}.");
            textOptions.Spacing = corrected;
        }
        else
        {
            Console.WriteLine($"Spacing {spacing} is within the allowed range.");
        }

        // Save the workbook
        workbook.Save("ValidatedSpacing.xlsx");
    }
}
