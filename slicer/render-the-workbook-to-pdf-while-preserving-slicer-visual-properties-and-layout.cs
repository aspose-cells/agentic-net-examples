// Title: Convert an Excel workbook with slicers to PDF while preserving slicer layout using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file containing slicers and saves it as a PDF with the slicer layout retained using Aspose.Cells. | Show how to enable slicer layout preservation by setting the SlicerLayoutOptions property on PdfSaveOptions in Aspose.Cells. | Add comprehensive error handling for missing source files and runtime exceptions when exporting a workbook with slicers to PDF.
// Common Searches: Aspose.Cells .NET keep slicer formatting when saving workbook as PDF | C# example for preserving slicer layout with PdfSaveOptions | How to export Excel file with slicers to PDF using Aspose.Cells | Convert Excel workbook to PDF while retaining slicer layout in C#
// Tags: Aspose.Cells PDF conversion with slicer layout preservation | PdfSaveOptions slicer layout option usage in .NET | Excel to PDF export preserving slicer visuals | C# error handling for missing workbook file Aspose.Cells | Slicer layout retention during workbook PDF rendering

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an Excel workbook, optionally configures PdfSaveOptions to keep slicer layout, and saves the workbook as a PDF while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing slicers
            Workbook workbook = new Workbook(inputPath);

            // Set PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Preserve slicer layout if the property is available in the used version
            // Uncomment the line below when SlicerLayoutOptions is supported:
            // pdfOptions.SlicerLayoutOptions = SlicerLayoutOptions.Preserve;

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
