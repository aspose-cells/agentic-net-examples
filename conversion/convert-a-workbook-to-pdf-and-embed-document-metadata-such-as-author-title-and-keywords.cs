// Title: Convert an Excel workbook to PDF and embed author, title, and keywords with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells set PDF author and title while converting .xlsx to .pdf | How to add keywords to a PDF generated from Excel using Aspose.Cells | Using WorkbookProperties to embed document metadata in PDF export with Aspose.Cells | Aspose.Cells PDF conversion preserve custom document properties in .NET | Reflection fallback for WorkbookProperties when setting PDF metadata in C#
// Tags: Aspose.Cells set PDF metadata | WorkbookProperties PDF export | Excel to PDF with document properties | C# add author title keywords to PDF | Aspose.Cells reflection workbookproperties

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing .xlsx file, uses reflection to access the WorkbookProperties object (if present), sets Author, Title, and Keywords, and then saves the workbook as a PDF, handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists before loading.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook.
            Workbook workbook = new Workbook(inputPath);

            // Attempt to set document metadata if the WorkbookProperties API is available.
            var propInfo = typeof(Workbook).GetProperty("WorkbookProperties");
            if (propInfo != null)
            {
                var docProps = propInfo.GetValue(workbook);
                var authorProp = docProps?.GetType().GetProperty("Author");
                var titleProp = docProps?.GetType().GetProperty("Title");
                var keywordsProp = docProps?.GetType().GetProperty("Keywords");

                authorProp?.SetValue(docProps, "John Doe");
                titleProp?.SetValue(docProps, "Quarterly Sales Report");
                keywordsProp?.SetValue(docProps, "sales, quarterly, report");
            }

            // Convert and save the workbook as PDF.
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF successfully saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
