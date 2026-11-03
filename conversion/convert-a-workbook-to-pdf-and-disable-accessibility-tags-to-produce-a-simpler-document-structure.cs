// Title: Convert an Excel workbook to PDF without accessibility tags using Aspose.Cells for .NET
// Common Searches: how to disable tagged PDF output when saving Excel as PDF with Aspose.Cells C# | Aspose.Cells convert .xlsx to PDF without accessibility tags | C# PdfSaveOptions remove accessibility tags from generated PDF | disable PDF accessibility in Aspose.Cells workbook conversion | set IsTaggedPdf false using Aspose.Cells PdfSaveOptions
// Tags: Aspose.Cells PdfSaveOptions turn off PDF tagging | Excel to PDF conversion without accessibility metadata | C# reflection set IsTaggedPdf property | PDF accessibility configuration Aspose.Cells | Workbook.Save PDF without tags

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Rendering; // Required for PdfSaveOptions

// The example loads an .xlsx workbook, creates a PdfSaveOptions object, uses reflection to access the AccessibilityOptions property and sets IsTaggedPdf to false, then saves the workbook as a PDF, producing a PDF that does not contain accessibility tags.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Attempt to disable tagged PDF generation if the API supports it
            PropertyInfo accessibilityProp = typeof(PdfSaveOptions).GetProperty("AccessibilityOptions");
            if (accessibilityProp != null)
            {
                // Create an instance of the AccessibilityOptions class
                object accessibilityInstance = Activator.CreateInstance(accessibilityProp.PropertyType);
                PropertyInfo isTaggedProp = accessibilityProp.PropertyType.GetProperty("IsTaggedPdf");
                if (isTaggedProp != null && isTaggedProp.CanWrite)
                {
                    // Set IsTaggedPdf to false to turn off accessibility tags
                    isTaggedProp.SetValue(accessibilityInstance, false);
                }
                // Assign the configured instance back to the PdfSaveOptions object
                accessibilityProp.SetValue(pdfOptions, accessibilityInstance);
            }

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
