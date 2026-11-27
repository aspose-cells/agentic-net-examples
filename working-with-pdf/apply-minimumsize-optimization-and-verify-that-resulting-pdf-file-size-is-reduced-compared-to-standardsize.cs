// Title: Generate a PDF from an Aspose.Cells workbook using MinimumSize optimization and compare its size to Standard optimization in C#
// Common Searches: C# Aspose.Cells save workbook as PDF with MinimumSize optimization | How to compare PDF file sizes between Standard and MinimumSize options in Aspose.Cells | Aspose.Cells PdfSaveOptions MinimumSize example for reducing PDF size | Verify that MinimumSize optimization produces a smaller PDF than Standard in .NET | Sample code to create a workbook and export to PDF using different optimization types
// Tags: PdfSaveOptions MinimumSize optimization | Aspose.Cells PDF size comparison | C# Aspose.Cells PDF generation Standard optimization | verify PDF file size reduction Aspose.Cells | minimum size PDF creation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook with sample data, saves it twice to PDF—once with the Standard optimization and once with MinimumSize optimization—then reads the file sizes and reports whether the MinimumSize PDF is smaller, demonstrating how to reduce PDF size using Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        try
        {
            // -----------------------------------------------------------------
            // 1. Create a sample workbook with enough data to generate a
            //    reasonably sized PDF file.
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the sheet with sample data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");

            for (int i = 2; i <= 101; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Product {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 2);
                sheet.Cells[$"C{i}"].PutValue(i * 1.5);
            }

            // -----------------------------------------------------------------
            // 2. Save the workbook to PDF using the Standard optimization.
            // -----------------------------------------------------------------
            string standardPdfPath = "StandardSize.pdf";
            PdfSaveOptions standardOptions = new PdfSaveOptions
            {
                // Use the default (balanced) optimization.
                OptimizationType = PdfOptimizationType.Standard
            };
            workbook.Save(standardPdfPath, standardOptions);

            // -----------------------------------------------------------------
            // 3. Save the same workbook to PDF using the MinimumSize optimization.
            // -----------------------------------------------------------------
            string minimumPdfPath = "MinimumSize.pdf";
            PdfSaveOptions minimumOptions = new PdfSaveOptions
            {
                // Force the smallest possible file size.
                OptimizationType = PdfOptimizationType.MinimumSize
            };
            workbook.Save(minimumPdfPath, minimumOptions);

            // -----------------------------------------------------------------
            // 4. Compare file sizes.
            // -----------------------------------------------------------------
            if (!File.Exists(standardPdfPath) || !File.Exists(minimumPdfPath))
                throw new FileNotFoundException("One or both PDF files were not created.");

            long standardSize = new FileInfo(standardPdfPath).Length;
            long minimumSize = new FileInfo(minimumPdfPath).Length;

            Console.WriteLine($"StandardSize PDF size : {standardSize:N0} bytes");
            Console.WriteLine($"MinimumSize PDF size : {minimumSize:N0} bytes");

            if (minimumSize < standardSize)
            {
                Console.WriteLine("Verification passed: MinimumSize optimization reduced the PDF file size.");
            }
            else
            {
                Console.WriteLine("Verification failed: MinimumSize optimization did NOT reduce the PDF file size.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
