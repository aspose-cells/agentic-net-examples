// Title: Save an Excel workbook with slicers to PDF, rendering slicer controls as static images using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, uses PdfSaveOptions to set SlicerRenderingMode to Image (when the property exists), and saves the workbook as a PDF. | Demonstrate how to use reflection in C# to detect the SlicerRenderingMode property on PdfSaveOptions before assigning the Image enum value. | Create a robust try‑catch block that validates the input workbook, configures PDF save options for slicer image rendering, and outputs a PDF with static slicer graphics.
// Common Searches: Aspose.Cells C# export Excel slicers to PDF as images | How to set PdfSaveOptions.SlicerRenderingMode to Image in .NET | Render slicer controls as static pictures when saving workbook to PDF with Aspose.Cells | Check for SlicerRenderingMode property at runtime before saving PDF | Convert Excel file with slicers to PDF using Aspose.Cells and keep slicers as images
// Tags: pdfsaveoptions slicer rendering mode | aspocells slicer pdf export | c# reflection pdfsaveoptions property | excel slicer static image pdf | convert workbook to pdf with slicer graphics

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads Input.xlsx, uses reflection to determine if PdfSaveOptions supports the SlicerRenderingMode property, sets it to the Image enum value when available, and saves the workbook as Output.pdf with all slicer controls rendered as static images.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.pdf";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // If the current Aspose.Cells version supports slicer rendering mode,
            // set it to render slicers as static images.
            var slicerProp = typeof(PdfSaveOptions).GetProperty("SlicerRenderingMode");
            if (slicerProp != null && slicerProp.CanWrite)
            {
                var enumType = slicerProp.PropertyType;
                var enumValue = Enum.Parse(enumType, "Image");
                slicerProp.SetValue(pdfOptions, enumValue);
            }

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
