// Title: Convert HTML to PDF and embed a custom ICC color profile with Aspose.Cells for .NET (C#)
// Common Searches: c# embed icc color profile when saving HTML workbook to PDF using Aspose.Cells | how to use PdfSaveOptions to add a custom ICC profile in Aspose.Cells conversion | Aspose.Cells HTML to PDF conversion with reflection for ColorProfile property | set PDF/A compliance and custom color profile in Aspose.Cells C# example | load HTML file into Aspose.Cells workbook and export to PDF with ICC profile
// Tags: Aspose.Cells HTML to PDF conversion with ICC profile | PdfSaveOptions embed custom color profile C# | reflection set ColorProfile property Aspose.Cells | custom ICC profile PDF/A generation using Aspose.Cells | load HTML into Workbook for PDF export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // For PdfSaveOptions

// The example loads an HTML file into an Aspose.Cells Workbook, configures PdfSaveOptions, and uses reflection to attach a custom ICC color profile (when supported) before saving the workbook as a PDF. It includes checks for file existence, version‑specific handling of the ColorProfile property, and optional PDF/A compliance settings.
class HtmlToPdfWithIcc
{
    static void Main()
    {
        // Paths (adjust as needed)
        string htmlPath = @"C:\Input\sample.html";
        string pdfPath = @"C:\Output\sample.pdf";
        string iccProfilePath = @"C:\Profiles\custom.icc";

        try
        {
            // Verify input HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"HTML file not found: {htmlPath}");
                return;
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(pdfPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Load the HTML into a Workbook
            Workbook workbook = new Workbook(htmlPath);

            // Prepare PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Embed custom ICC profile if the API supports it (available in newer versions)
            // In older versions PdfSaveOptions does not expose a ColorProfile property.
            // The following block uses reflection to set it when present, otherwise skips it.
            if (File.Exists(iccProfilePath))
            {
                try
                {
                    var colorProfileType = typeof(PdfSaveOptions).Assembly.GetType("Aspose.Cells.Rendering.ColorProfile");
                    if (colorProfileType != null)
                    {
                        // Create an instance of ColorProfile using the file path constructor
                        var ctor = colorProfileType.GetConstructor(new[] { typeof(string) });
                        if (ctor != null)
                        {
                            var colorProfileInstance = ctor.Invoke(new object[] { iccProfilePath });
                            var prop = typeof(PdfSaveOptions).GetProperty("ColorProfile");
                            if (prop != null && prop.CanWrite)
                            {
                                prop.SetValue(pdfOptions, colorProfileInstance);
                            }
                            else
                            {
                                Console.WriteLine("PdfSaveOptions does not have a writable ColorProfile property. ICC profile will not be embedded.");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Unable to locate a suitable constructor for ColorProfile. ICC profile will not be embedded.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("ColorProfile type not found in the current Aspose.Cells version. ICC profile will not be embedded.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to embed ICC profile: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("ICC profile not found. PDF will be saved without a custom profile.");
            }

            // Optional: set PDF compliance (e.g., PDF/A-1b)
            // pdfOptions.Compliance = PdfCompliance.PdfA1b;

            // Save as PDF
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine("HTML has been successfully converted to PDF.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
