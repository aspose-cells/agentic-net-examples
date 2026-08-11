// Title: C# – Convert Excel WordArt with Gradient Fill to PDF using Aspose.Cells
// Description: Loads an Excel workbook, finds WordArt shapes, changes their fill to a two‑color gradient, saves a temporary XLSX, and converts it to PDF with Aspose.Cells while cleaning up the temporary file.
// Keywords: Aspose.Cells WordArt gradient | C# convert Excel to PDF | WordArt gradient fill Aspose | Excel shape fill type gradient | ConversionUtility PDF export | .NET Excel to PDF gradient
// Common Searches: how to keep WordArt gradient when converting Excel to PDF with Aspose.Cells | set gradient fill for WordArt shapes before PDF export C# | Aspose.Cells convert Excel WordArt to PDF with accurate colors | apply two‑color diagonal gradient to WordArt in .NET
// Developer Intent: Apply a gradient fill to WordArt objects in an Excel file and export the workbook to PDF while preserving the gradient appearance.
// Use Cases: Batch‑process Excel templates that contain WordArt, applying a brand‑specific gradient before generating PDF reports. | Dynamically adjust WordArt gradient colors based on runtime data (e.g., status indicators) and create a PDF snapshot for distribution. | Produce a PDF catalog from an Excel design where WordArt titles must use a diagonal red‑to‑blue gradient for consistent branding.
// AI Prompts: Show how to use a three‑color gradient on WordArt shapes before PDF conversion with Aspose.Cells. | Provide an example that converts an Excel file with WordArt to PDF without creating a temporary XLSX file. | Explain how to retain WordArt outline, shadow, and other formatting when exporting to PDF using Aspose.Cells.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Utility;

// Loads an Excel workbook, finds WordArt shapes, changes their fill to a two‑color gradient, saves a temporary XLSX, and converts it to PDF with Aspose.Cells while cleaning up the temporary file.
public class WordArtPdfConverter
{
    /// <param name="inputExcelPath">Full path to the source Excel file.</param>
    /// <param name="outputPdfPath">Full path where the resulting PDF will be saved.</param>
    public static void Convert(string inputExcelPath, string outputPdfPath)
    {
        // Verify that the source Excel file exists.
        if (!File.Exists(inputExcelPath))
        {
            Console.WriteLine($"Error: Input file not found – {inputExcelPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook from the specified Excel file.
            workbook = new Workbook(inputExcelPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Iterate through all worksheets in the workbook.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all shapes on the worksheet.
            foreach (Shape shape in sheet.Shapes)
            {
                // Process only WordArt shapes using the IsWordArt property.
                if (shape.IsWordArt)
                {
                    // Set the fill type to Gradient to enable gradient operations.
                    shape.Fill.FillType = FillType.Gradient;

                    // Obtain the GradientFill object; it may be null if the fill type is not gradient.
                    GradientFill gradientFill = shape.Fill.GradientFill;
                    if (gradientFill != null)
                    {
                        // Apply a two‑color gradient (example: Red → Blue, diagonal down, variant 1).
                        gradientFill.SetTwoColorGradient(
                            Color.Red,               // First gradient color
                            Color.Blue,              // Second gradient color
                            GradientStyleType.DiagonalDown,
                            1);                      // Variant (1‑4)
                    }
                }
            }
        }

        // Save the modified workbook to a temporary XLSX file.
        string tempXlsxPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
        try
        {
            workbook.Save(tempXlsxPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save temporary workbook: {ex.Message}");
            return;
        }

        try
        {
            // Convert the temporary XLSX file to PDF using the ConversionUtility rule.
            ConversionUtility.Convert(tempXlsxPath, outputPdfPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Conversion failed: {ex.Message}");
        }
        finally
        {
            // Clean up the temporary file.
            if (File.Exists(tempXlsxPath))
            {
                try
                {
                    File.Delete(tempXlsxPath);
                }
                catch
                {
                    // Ignored – best‑effort cleanup.
                }
            }
        }
    }

    // Example usage
    public static void Main()
    {
        string sourceExcel = "WordArtSample.xlsx";   // Replace with your source file path
        string targetPdf   = "WordArtOutput.pdf";    // Desired PDF output path

        try
        {
            Convert(sourceExcel, targetPdf);
            Console.WriteLine("Conversion completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
