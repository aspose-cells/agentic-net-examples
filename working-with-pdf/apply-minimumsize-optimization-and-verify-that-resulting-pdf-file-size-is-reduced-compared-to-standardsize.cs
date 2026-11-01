// Title: Save Excel workbook as PDF with MinimumSize optimization using Aspose.Cells for .NET and compare file size to StandardSize
// Common Searches: Aspose.Cells C# how to enable MinimumSize PDF optimization when saving a workbook | compare PDF output size using StandardSize vs MinimumSize in Aspose.Cells .NET | set PdfSaveOptions OptimizationOptions property via reflection in C# Aspose.Cells | reduce Excel to PDF file size with MinimumSize option in Aspose.Cells | verify PDF size reduction after applying MinimumSize optimization in .NET
// Tags: Aspose.Cells PDF save with MinimumSize setting | Reflection set PdfSaveOptions OptimizationOptions C# | Excel workbook PDF size reduction Aspose.Cells | StandardSize PDF output Aspose.Cells comparison | File size measurement Aspose.Cells PDF export

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads or creates an Excel workbook, saves it to PDF twice—first with the default StandardSize settings and then with MinimumSize optimization by configuring PdfSaveOptions.OptimizationOptions via reflection, measures both file sizes, and reports whether the MinimumSize option produced a smaller PDF.
class MinimumSizeOptimizationDemo
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string excelPath = "input.xlsx";

            // Load the workbook; create a simple one if the file does not exist
            Workbook workbook;
            if (File.Exists(excelPath))
            {
                workbook = new Workbook(excelPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Cells["A1"].PutValue("Sample data");
            }

            // Save PDF with default (StandardSize) optimization
            string standardPdfPath = "standard.pdf";
            workbook.Save(standardPdfPath, SaveFormat.Pdf);

            // Save PDF with MinimumSize optimization (if supported)
            string minimumPdfPath = "minimum.pdf";
            PdfSaveOptions minSizeOptions = new PdfSaveOptions();

            // Attempt to set the OptimizationOptions property via reflection
            PropertyInfo optProp = typeof(PdfSaveOptions).GetProperty("OptimizationOptions");
            if (optProp != null && optProp.CanWrite)
            {
                Type enumType = optProp.PropertyType; // Expected: PdfOptimizationOptions
                object minSizeEnum = Enum.Parse(enumType, "MinimumSize");
                optProp.SetValue(minSizeOptions, minSizeEnum);
            }

            workbook.Save(minimumPdfPath, minSizeOptions);

            // Get file sizes
            long standardSize = new FileInfo(standardPdfPath).Length;
            long minimumSize = new FileInfo(minimumPdfPath).Length;

            // Output results
            Console.WriteLine($"Standard PDF size: {standardSize} bytes");
            Console.WriteLine($"MinimumSize PDF size: {minimumSize} bytes");

            if (minimumSize < standardSize)
            {
                Console.WriteLine("Success: MinimumSize optimization reduced the PDF file size.");
            }
            else
            {
                Console.WriteLine("Warning: MinimumSize optimization did not reduce the PDF file size.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
