// Title: Export Excel with Smart Markers, Charts & Images to PDF using Aspose.Cells for .NET
// Description: Load an Excel template that contains smart markers, charts, and embedded images, bind a collection of Product objects, process the markers, and save the workbook as a PDF. PdfSaveOptions with ExportDocumentStructure and CalculateFormula ensure that charts, graphics, and calculated values are retained in the PDF output.
// Keywords: Aspose.Cells PDF export | smart markers to PDF | preserve charts Aspose.Cells | export images to PDF .NET | PdfSaveOptions ExportDocumentStructure | CalculateFormula PDF | C# Excel to PDF conversion | Aspose.Cells WorkbookDesigner | Excel template PDF generation
// Common Searches: How to export an Excel file with smart markers and charts to PDF using Aspose.Cells | Aspose.Cells preserve images and formulas when saving as PDF | C# code for PDF export with smart markers | PdfSaveOptions settings for chart retention | Export Excel workbook to PDF with Aspose.Cells .NET
// Developer Intent: Create a PDF from a smart‑marker populated workbook that keeps all charts, images, and calculated formulas intact.
// Use Cases: Generate product catalogs by filling an Excel template with product names, prices, and photos, then export the result to a printable PDF. | Automate reporting where charts are driven by smart markers and must appear unchanged in the final PDF document. | Integrate PDF generation into a .NET application while ensuring formulas are recalculated and document structure is preserved.
// AI Prompts: Write C# code that loads an Excel template with smart markers, binds a list of objects containing image byte arrays, processes the markers, and saves the workbook as a PDF preserving charts and formulas using Aspose.Cells. | Explain the impact of PdfSaveOptions properties ExportDocumentStructure and CalculateFormula on PDF output for workbooks that contain smart markers and charts. | Provide troubleshooting steps for missing images or incorrect chart data when exporting a smart‑marker populated workbook to PDF with Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    // Sample data class for smart markers
    // Load an Excel template that contains smart markers, charts, and embedded images, bind a collection of Product objects, process the markers, and save the workbook as a PDF. PdfSaveOptions with ExportDocumentStructure and CalculateFormula ensure that charts, graphics, and calculated values are retained in the PDF output.
    public class Product
    {
        public string? Name { get; set; }
        public double Price { get; set; }
        public byte[]? Image { get; set; }
    }

    public class ExportWorkbookToPdf
    {
        public static void Run()
        {
            try
            {
                const string templatePath = "Template.xlsx";
                const string outputPath = "Output.pdf";

                // Verify template file exists
                if (!File.Exists(templatePath))
                    throw new FileNotFoundException($"Template file not found: {templatePath}");

                // Load the Excel template that contains smart markers and charts
                Workbook workbook = new Workbook(templatePath);

                // Prepare data source for smart markers
                List<Product> products = new List<Product>
                {
                    CreateProduct("Apple", "apple.jpg"),
                    CreateProduct("Banana", "banana.jpg"),
                    CreateProduct("Cherry", "cherry.jpg")
                };

                // Initialize WorkbookDesigner, bind data source and process smart markers
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("Products", products);
                designer.Process(); // populates smart markers, updates charts/graphics

                // Configure PDF save options to preserve document structure and calculate formulas
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    ExportDocumentStructure = true,
                    CalculateFormula = true
                };

                // Save the populated workbook as PDF (charts and graphics are retained)
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during PDF export: {ex.Message}");
            }
        }

        // Helper method to create a Product instance with image validation
        private static Product CreateProduct(string name, string imageFileName)
        {
            if (!File.Exists(imageFileName))
                throw new FileNotFoundException($"Image file not found: {imageFileName}");

            return new Product
            {
                Name = name,
                Price = GetPriceForProduct(name),
                Image = File.ReadAllBytes(imageFileName)
            };
        }

        // Simple price lookup (could be replaced with real logic)
        private static double GetPriceForProduct(string name) => name switch
        {
            "Apple" => 1.20,
            "Banana" => 0.80,
            "Cherry" => 2.50,
            _ => 0.0
        };
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportWorkbookToPdf.Run();
        }
    }
}
