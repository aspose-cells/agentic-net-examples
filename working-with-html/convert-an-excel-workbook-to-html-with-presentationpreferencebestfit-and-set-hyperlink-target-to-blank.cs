// Title: Convert an Excel workbook to HTML with BestFit layout and _blank hyperlink target in C# using Aspose.Cells
// Common Searches: C# Aspose.Cells export Excel to HTML with BestFit presentation preference | Set hyperlink target to _blank when saving workbook as HTML using Aspose.Cells | How to use HtmlSaveOptions to fit column widths and open links in a new tab in .NET | Aspose.Cells HTML conversion preserve column width best fit and blank link target | Reflection workaround for HtmlSaveOptions properties in older Aspose.Cells versions
// Tags: Aspose.Cells HtmlSaveOptions BestFit layout | Aspose.Cells HTML export hyperlink target blank | C# Excel to HTML conversion with Aspose.Cells | HtmlSaveOptions PresentationPreference BestFit | HtmlSaveOptions HyperlinkTarget _blank reflection

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel file, configures HtmlSaveOptions to use the BestFit presentation preference and sets hyperlink targets to "_blank" (using reflection for version compatibility), then saves the workbook as an HTML file while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Set PresentationPreference to BestFit if the property exists (for compatibility with different versions)
            var prefProp = typeof(HtmlSaveOptions).GetProperty("PresentationPreference");
            if (prefProp != null && prefProp.CanWrite)
            {
                var enumType = prefProp.PropertyType;
                var bestFitValue = Enum.Parse(enumType, "BestFit");
                prefProp.SetValue(htmlOptions, bestFitValue);
            }

            // Set hyperlink target to open in a new window/tab if the property exists
            var hyperlinkTargetProp = typeof(HtmlSaveOptions).GetProperty("HyperlinkTarget");
            if (hyperlinkTargetProp != null && hyperlinkTargetProp.CanWrite)
            {
                hyperlinkTargetProp.SetValue(htmlOptions, "_blank");
            }

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
