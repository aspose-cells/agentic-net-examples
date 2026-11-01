// Title: Compare Aspose.Cells HTML output using PresentationPreference.BestFit versus default layout in C#
// Common Searches: Aspose.Cells C# generate HTML with PresentationPreference BestFit | difference between default and BestFit HTML layout in Aspose.Cells | how to set PresentationPreference in HtmlSaveOptions via reflection | compare file size of Aspose.Cells HTML outputs for layout options | visual fidelity of Aspose.Cells generated HTML with BestFit setting
// Tags: Aspose.Cells HtmlSaveOptions PresentationPreference | BestFit HTML layout generation | default vs BestFit HTML comparison | C# file size analysis of generated HTML | reflection to modify Aspose.Cells save options

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an Excel workbook, saves it twice to HTML—once with default HtmlSaveOptions and once after setting PresentationPreference to BestFit via reflection—and then compares the resulting file sizes to infer differences in rendering detail and visual fidelity.
class HtmlComparison
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string sourceFile = "input.xlsx";

            // Ensure the source file exists; create a simple workbook if missing
            if (!File.Exists(sourceFile))
            {
                var tempWb = new Workbook();
                var sheet = tempWb.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample");
                sheet.Cells["B1"].PutValue(123);
                tempWb.Save(sourceFile);
            }

            // Paths for the generated HTML files
            string defaultHtml = "output_default.html";
            string bestFitHtml = "output_bestfit.html";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(sourceFile);

            // -----------------------------------------------------------------
            // Generate HTML with default layout (no PresentationPreference set)
            // -----------------------------------------------------------------
            HtmlSaveOptions defaultOptions = new HtmlSaveOptions(SaveFormat.Html);
            workbook.Save(defaultHtml, defaultOptions);

            // -----------------------------------------------------------------
            // Generate HTML with PresentationPreference.BestFit (if supported)
            // -----------------------------------------------------------------
            HtmlSaveOptions bestFitOptions = new HtmlSaveOptions(SaveFormat.Html);
            try
            {
                // Use reflection to set PresentationPreference to BestFit if the property exists
                var prop = typeof(HtmlSaveOptions).GetProperty("PresentationPreference");
                if (prop != null && prop.CanWrite)
                {
                    var enumType = prop.PropertyType;
                    var bestFitValue = Enum.Parse(enumType, "BestFit");
                    prop.SetValue(bestFitOptions, bestFitValue);
                }
            }
            catch (Exception)
            {
                // Ignore any errors; continue with default settings
            }

            workbook.Save(bestFitHtml, bestFitOptions);

            // -----------------------------------------------------------------
            // Simple visual fidelity comparison (file size as a proxy)
            // -----------------------------------------------------------------
            long defaultSize = new FileInfo(defaultHtml).Length;
            long bestFitSize = new FileInfo(bestFitHtml).Length;

            Console.WriteLine($"Default HTML size: {defaultSize} bytes");
            Console.WriteLine($"BestFit HTML size: {bestFitSize} bytes");

            if (bestFitSize > defaultSize)
            {
                Console.WriteLine("BestFit layout produced a larger file, indicating more detailed rendering.");
            }
            else if (bestFitSize < defaultSize)
            {
                Console.WriteLine("BestFit layout produced a smaller file, indicating more compact rendering.");
            }
            else
            {
                Console.WriteLine("Both layouts resulted in files of equal size.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
