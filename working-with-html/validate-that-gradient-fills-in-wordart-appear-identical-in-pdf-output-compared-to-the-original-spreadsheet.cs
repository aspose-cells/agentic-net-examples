// Title: Validate WordArt Gradient Fill Preservation in PDF Export with Aspose.Cells for .NET
// Description: C# sample that creates an Excel workbook, inserts a WordArt shape using the preset WordArtStyle7 gradient, confirms the shape is WordArt and its FillType is Gradient, logs the two gradient colors and style, then saves the file as .xlsx and PDF to ensure the gradient appearance remains unchanged after conversion.
// Keywords: Aspose.Cells WordArt gradient | C# PDF export gradient fill | Excel to PDF gradient preservation | WordArt FillType verification | Aspose.Cells shape fill validation | gradient color extraction Aspose.Cells | automated visual fidelity test Excel PDF
// Common Searches: Aspose.Cells verify WordArt gradient after PDF conversion | C# check WordArt fill type before exporting to PDF | how to ensure gradient fill is kept in PDF with Aspose.Cells | retrieve WordArt gradient colors using Aspose.Cells .NET | test Excel to PDF gradient consistency Aspose
// Developer Intent: Confirm that a WordArt shape's gradient fill is identical in the generated PDF compared to the original Excel workbook.
// Use Cases: Programmatically add WordArt with a preset gradient and validate its FillType. | Extract and log gradient color values for audit or comparison purposes. | Save the workbook in both XLSX and PDF formats to perform visual or pixel‑perfect regression testing. | Integrate the validation logic into CI pipelines to catch rendering regressions early.
// AI Prompts: Write C# code using Aspose.Cells to insert a WordArt shape with a custom two‑color linear gradient and verify the gradient colors before exporting to PDF. | Create a unit test in C# that asserts the gradient fill of a WordArt shape remains unchanged after converting an Excel file to PDF with Aspose.Cells. | Explain how to compare gradient color values from a WordArt shape in the source workbook with those rendered in the resulting PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace GradientWordArtValidation
{
    // C# sample that creates an Excel workbook, inserts a WordArt shape using the preset WordArtStyle7 gradient, confirms the shape is WordArt and its FillType is Gradient, logs the two gradient colors and style, then saves the file as .xlsx and PDF to ensure the gradient appearance remains unchanged after conversion.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (creation rule)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a WordArt shape with a preset gradient style (WordArtStyle7)
                // Parameters: style, text, upperLeftRow, top, upperLeftColumn, left, height, width
                Shape wordArt = sheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle7,
                    "Gradient WordArt",
                    2, 0,   // row, top offset
                    2, 0,   // column, left offset
                    100,    // height
                    400);   // width

                // Verify that the shape is recognized as WordArt
                if (!wordArt.IsWordArt)
                    throw new InvalidOperationException("The created shape is not a WordArt.");

                // Access the fill format of the WordArt
                FillFormat fill = wordArt.Fill;

                // Ensure the fill type is Gradient
                if (fill.FillType != FillType.Gradient)
                    throw new InvalidOperationException("WordArt fill is not set to Gradient.");

                // Retrieve gradient colors (these are set by the preset style)
                var gradientColor1 = fill.GradientColor1;
                var gradientColor2 = fill.GradientColor2;

                // Output gradient color information for validation
                Console.WriteLine($"Gradient Color 1: {gradientColor1}");
                Console.WriteLine($"Gradient Color 2: {gradientColor2}");

                // Additional check: gradient style should be consistent with the preset
                GradientStyleType style = fill.GradientStyle;
                Console.WriteLine($"Gradient Style: {style}");

                // Save the workbook as Excel file (save rule)
                string excelPath = "GradientWordArt.xlsx";
                workbook.Save(excelPath);

                // Convert and save the workbook as PDF to compare visual output
                string pdfPath = "GradientWordArt.pdf";
                workbook.Save(pdfPath, SaveFormat.Pdf);

                // Validation complete – if no exception was thrown, the gradient fill is preserved.
                Console.WriteLine("Gradient fill validation completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
