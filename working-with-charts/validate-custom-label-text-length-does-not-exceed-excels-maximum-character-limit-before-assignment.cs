// Title: C# – Validate and Truncate ActiveX Label Caption to Excel’s 32,767‑Character Limit with Aspose.Cells
// Description: Shows how to create a workbook, turn on Excel restriction checking, add a Label ActiveX control, compare its Caption to Excel’s 32,767‑character maximum, truncate the string when necessary, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# ActiveX label | Excel caption length limit | 32,767 characters | CheckExcelRestriction | truncate label text | validate ActiveX control text | Excel restriction checking | ActiveX label caption | Aspose.Cells .NET
// Common Searches: Aspose.Cells truncate ActiveX label caption | C# check Excel character limit for label | enable CheckExcelRestriction in Aspose.Cells | validate label text length before saving workbook | Excel maximum characters for ActiveX label control
// Developer Intent: Ensure an ActiveX label’s Caption stays within Excel’s 32,767‑character limit to avoid errors.
// Use Cases: Add a Label ActiveX control to a worksheet and safely assign a long string by enforcing the Excel character cap. | Activate workbook.Settings.CheckExcelRestriction so Aspose.Cells automatically respects Excel limits during processing. | Save a workbook after truncating oversized label text, preventing runtime exceptions.
// AI Prompts: Generate C# code with Aspose.Cells that adds an ActiveX label and automatically shortens its Caption if it exceeds 32,767 characters. | Explain how to enable CheckExcelRestriction in Aspose.Cells and programmatically verify a string before setting it as a label’s Caption. | Provide an example of handling oversized label text in Aspose.Cells, including logging the truncation and continuing the save operation.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, turn on Excel restriction checking, add a Label ActiveX control, compare its Caption to Excel’s 32,767‑character maximum, truncate the string when necessary, and save the file using Aspose.Cells for .NET.
    public class ValidateLabelTextLength
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Ensure Excel restriction checking is enabled (default is true)
                workbook.Settings.CheckExcelRestriction = true;

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a Label ActiveX control to the worksheet
                // Parameters: control type, upper left row, upper left column, top offset, left offset, height, width
                Shape shape = worksheet.Shapes.AddActiveXControl(
                    ControlType.Label, 2, 2, 0, 0, 100, 200);

                // Cast the ActiveXControl to LabelActiveXControl
                LabelActiveXControl label = (LabelActiveXControl)shape.ActiveXControl;

                // Example text to assign to the label (33,000 characters, exceeds Excel limit)
                string labelText = new string('X', 33000);

                // Excel's maximum allowed characters for a cell/label text is 32,767
                const int ExcelMaxTextLength = 32767;

                // Validate length before assignment
                if (labelText.Length > ExcelMaxTextLength)
                {
                    // Truncate the text to the maximum allowed length
                    labelText = labelText.Substring(0, ExcelMaxTextLength);
                }

                // Assign the validated (or truncated) text to the label's caption
                label.Caption = labelText;

                // Save the workbook
                workbook.Save("ValidatedLabelText.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ValidateLabelTextLength.Run();
        }
    }
}
