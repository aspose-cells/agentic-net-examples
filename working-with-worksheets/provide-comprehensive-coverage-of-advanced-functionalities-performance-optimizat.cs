using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsAdvancedDemo
{
    // Custom function definition that marks the first parameter to be evaluated in array mode
    public class MyCustomFunctionDefinition : CustomFunctionDefinition
    {
        public override int[] GetArrayModeParameters(string functionName)
        {
            // For demonstration, treat the first parameter as array mode for any custom function
            return new int[] { 0 };
        }
    }

    // Custom calculation engine that implements a simple sum of two parameters
    public class MyCustomCalculationEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Example custom function name
            if (data.FunctionName.Equals("MYSUM", StringComparison.OrdinalIgnoreCase))
            {
                // Retrieve first and second parameters (could be scalar or ReferredArea)
                object param1 = data.GetParamValue(0);
                object param2 = data.GetParamValue(1);

                double val1 = Convert.ToDouble(ExtractValue(param1));
                double val2 = Convert.ToDouble(ExtractValue(param2));

                // Set the calculated result
                data.CalculatedValue = val1 + val2;
            }
        }

        // Helper to extract a scalar value from possible parameter types
        private object ExtractValue(object param)
        {
            if (param is ReferredArea area)
            {
                // Get the top‑left cell value of the area
                return area.GetValue(0, 0);
            }
            return param;
        }
    }

    // Stream provider that writes linked resources (e.g., images) to a custom directory
    public class ExportStreamProvider : IStreamProvider
    {
        private readonly string _outputDir;

        public ExportStreamProvider(string outputDir)
        {
            _outputDir = outputDir;
        }

        public void InitStream(StreamProviderOptions options)
        {
            // Ensure the output directory exists
            Directory.CreateDirectory(_outputDir);

            // Determine the target file name
            string fileName = Path.GetFileName(options.DefaultPath ?? "resource.bin");
            string fullPath = Path.Combine(_outputDir, fileName);

            // Create the stream for writing
            options.Stream = new FileStream(fullPath, FileMode.Create, FileAccess.Write);
            options.CustomPath = fullPath;
        }

        public void CloseStream(StreamProviderOptions options)
        {
            options.Stream?.Close();
            options.Stream = null;
        }
    }

    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook and populate data
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill a 10x5 range with sample text
            for (int i = 0; i < 10; i++)
                for (int j = 0; j < 5; j++)
                    sheet.Cells[i, j].PutValue($"R{i + 1}C{j + 1}");

            // Add a numeric column for custom calculations
            sheet.Cells["F1"].PutValue("Value1");
            sheet.Cells["G1"].PutValue("Value2");
            sheet.Cells["F2"].PutValue(12);
            sheet.Cells["G2"].PutValue(34);
            sheet.Cells["F3"].PutValue(56);
            sheet.Cells["G3"].PutValue(78);

            // -------------------------------------------------
            // 2. Performance optimization with AccessCacheOptions
            // -------------------------------------------------
            workbook.StartAccessCache(AccessCacheOptions.PositionAndSize | AccessCacheOptions.CellsData);
            // Simulate intensive read operations
            for (int i = 0; i < 1000; i++)
            {
                var val = sheet.Cells[i % 10, i % 5].StringValue;
            }
            workbook.CloseAccessCache(AccessCacheOptions.PositionAndSize | AccessCacheOptions.CellsData);

            // -------------------------------------------------
            // 3. Insert an image that will be reused
            // -------------------------------------------------
            // Ensure the image file exists at the specified path
            string imagePath = "logo.png";
            if (File.Exists(imagePath))
            {
                int picIdx = sheet.Pictures.Add(1, 1, imagePath);
                Picture pic = sheet.Pictures[picIdx];
                pic.Width = 100;
                pic.Height = 100;

                // Duplicate the same image in another cell
                int picIdx2 = sheet.Pictures.Add(5, 3, imagePath);
                Picture pic2 = sheet.Pictures[picIdx2];
                pic2.Width = 100;
                pic2.Height = 100;
            }

            // -------------------------------------------------
            // 4. Register and use a custom function
            // -------------------------------------------------
            // Update the workbook with the custom function definition
            workbook.UpdateCustomFunctionDefinition(new MyCustomFunctionDefinition());

            // Set a formula that uses the custom function MYSUM
            sheet.Cells["H2"].Formula = "=MYSUM(F2, G2)";
            sheet.Cells["H3"].Formula = "=MYSUM(F3, G3)";

            // Configure calculation options with the custom engine
            CalculationOptions calcOptions = new CalculationOptions
            {
                CustomEngine = new MyCustomCalculationEngine(),
                IgnoreError = false,
                Recursive = true
            };

            // Perform calculation
            workbook.CalculateFormula(calcOptions);

            // -------------------------------------------------
            // 5. Save as HTML with CSS custom properties optimization
            // -------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                EnableCssCustomProperties = true,          // Optimize repeated base64 images
                ExportActiveWorksheetOnly = false,         // Export all worksheets
                ExportImagesAsBase64 = true,               // Embed images as base64
                CssStyles = "body {font-family:Arial;}",   // Additional CSS
                SaveAsSingleFile = true                    // Single HTML file
            };
            workbook.Save("AdvancedDemo_Optimized.html", htmlOptions);

            // -------------------------------------------------
            // 6. Save as PDF with minimum size optimization
            // -------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OptimizationType = PdfOptimizationType.MinimumSize
            };
            workbook.Save("AdvancedDemo_MinSize.pdf", pdfOptions);

            // -------------------------------------------------
            // 7. Save as DOCX (editable tables)
            // -------------------------------------------------
            DocxSaveOptions docxOptions = new DocxSaveOptions();
            workbook.Save("AdvancedDemo.docx", docxOptions);

            // -------------------------------------------------
            // 8. Render the first worksheet to a PNG image with optimization
            // -------------------------------------------------
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                OnePagePerSheet = true,
                IsOptimized = true                         // Optimize borders and file size
            };
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            using (MemoryStream imgStream = new MemoryStream())
            {
                renderer.ToImage(0, imgStream);
                File.WriteAllBytes("AdvancedDemo.png", imgStream.ToArray());
            }

            // -------------------------------------------------
            // 9. Use a custom stream provider for linked resources
            // -------------------------------------------------
            workbook.Settings.ResourceProvider = new ExportStreamProvider("CustomResources");
            // Save again to demonstrate the custom provider (HTML output)
            HtmlSaveOptions htmlCustomRes = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = false,   // Force external image files
                AttachedFilesDirectory = "CustomResources"
            };
            workbook.Save("AdvancedDemo_WithCustomResources.html", htmlCustomRes);

            // -------------------------------------------------
            // 10. Demonstrate CellsHelper utilities (optional)
            // -------------------------------------------------
            string safeName = CellsHelper.CreateSafeSheetName("Invalid/Name*With:Chars");
            Console.WriteLine($"Safe sheet name generated: {safeName}");

            // -------------------------------------------------
            // Completion message
            // -------------------------------------------------
            Console.WriteLine("All advanced demos executed successfully.");
        }
    }
}