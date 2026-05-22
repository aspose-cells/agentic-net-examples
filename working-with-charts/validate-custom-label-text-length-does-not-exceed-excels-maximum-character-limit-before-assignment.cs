using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsLabelValidation
{
    class Program
    {
        // Excel's maximum allowed characters for a text label (same as cell limit)
        const int MaxLabelLength = 32767;

        // Validates the label text and truncates it if it exceeds the limit
        static string ValidateLabelText(string text)
        {
            if (text == null) return string.Empty;
            return text.Length <= MaxLabelLength ? text : text.Substring(0, MaxLabelLength);
        }

        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Ensure Excel restriction checking is enabled (optional, throws if over limit)
            workbook.Settings.CheckExcelRestriction = true;

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a label shape to the worksheet
            // Parameters: upper left row, upper left column, lower right row, lower right column, width, height
            Shape labelShape = sheet.Shapes.AddLabel(1, 1, 5, 5, 200, 50);

            // Example of a long custom text (may exceed Excel's limit)
            string longText = new string('X', 35000); // 35,000 characters

            // Validate and possibly truncate the text before assignment
            string safeText = ValidateLabelText(longText);

            // Assign the validated text to the label
            labelShape.Text = safeText;

            // Optionally, display the final length for verification
            Console.WriteLine("Label text length after validation: " + safeText.Length);

            // Save the workbook
            workbook.Save("LabelWithValidatedText.xlsx");
        }
    }
}