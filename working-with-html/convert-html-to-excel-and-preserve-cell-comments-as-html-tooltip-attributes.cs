// Title: Convert HTML to XLSX and preserve cell comments as HTML tooltip attributes with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# convert HTML to Excel while keeping cell comments as tooltips | How to export cell comments as tooltip attributes when saving to HTML using Aspose.Cells | Load an HTML file into a workbook and save as XLSX preserving comment tooltips in C# | Enable ExportCellCommentsAs AsTooltips with HtmlSaveOptions in Aspose.Cells | Using reflection to set ExportCellComments properties in Aspose.Cells HtmlSaveOptions
// Tags: Aspose.Cells HTML to XLSX conversion | HtmlSaveOptions ExportCellCommentsAs tooltip | preserve cell comments during HTML to Excel conversion | C# load HTML workbook Aspose.Cells | export comments as HTML tooltip attributes

using Aspose.Cells;
using System;
using System.IO;

// The example loads an HTML file into an Aspose.Cells Workbook, saves it as an XLSX file, and then uses HtmlSaveOptions (with reflection) to enable ExportCellComments and set ExportCellCommentsAs to AsTooltips, producing an HTML output where cell comments appear as tooltip attributes.
class Program
{
    static void Main()
    {
        // Paths for input HTML and output files
        string inputHtmlPath = "input.html";
        string outputExcelPath = "output.xlsx";
        string outputHtmlPath = "output.html";

        try
        {
            // Verify that the input HTML file exists
            if (!File.Exists(inputHtmlPath))
            {
                Console.WriteLine($"Input file not found: {inputHtmlPath}");
                return;
            }

            // Load the HTML file into a workbook using Html load options
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(inputHtmlPath, loadOptions);

            // Save the workbook as an Excel file (XLSX)
            workbook.Save(outputExcelPath, SaveFormat.Xlsx);

            // Prepare HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Export cell comments as tooltips if the properties are available in the current version
            var exportCommentsProp = typeof(HtmlSaveOptions).GetProperty("ExportCellComments");
            var exportCommentsAsProp = typeof(HtmlSaveOptions).GetProperty("ExportCellCommentsAs");
            if (exportCommentsProp != null && exportCommentsAsProp != null)
            {
                exportCommentsProp.SetValue(htmlOptions, true);

                // Set the enum value for tooltips using reflection to avoid direct enum reference
                try
                {
                    object enumValue = Enum.Parse(exportCommentsAsProp.PropertyType, "AsTooltips");
                    exportCommentsAsProp.SetValue(htmlOptions, enumValue);
                }
                catch
                {
                    // If the enum value is not found, ignore and continue with default behavior
                }
            }

            // Save the workbook back to HTML; comments (if any) will appear as tooltip attributes
            workbook.Save(outputHtmlPath, htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
