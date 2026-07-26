// Title: Export Excel to HTML with fallback border styles using HtmlSaveOptions.ExportSimilarBorderStyle in Aspose.Cells for .NET
// Description: Shows how to configure Aspose.Cells for .NET to save a workbook as HTML with ExportSimilarBorderStyle, IE compatibility, and collapsed borders, so unsupported border types (e.g., double) fall back to a compatible style in legacy browsers.
// Keywords: Aspose.Cells | HtmlSaveOptions | ExportSimilarBorderStyle | fallback border | legacy browser HTML export | IsIECompatible | IsBorderCollapsed | C# Excel to HTML | .NET HTML export | border style compatibility
// Common Searches: Aspose.Cells ExportSimilarBorderStyle example | HTML export with fallback borders in Aspose.Cells | Enable IE compatibility when saving Excel as HTML | How to collapse table borders in Aspose.Cells HTML output | Double border fallback in legacy browsers Aspose
// Developer Intent: Generate HTML from an Excel workbook that automatically substitutes unsupported border styles with a similar, widely supported style and ensures the markup renders correctly in older browsers.
// Use Cases: Create web‑ready reports where double borders degrade to thin borders for Internet Explorer 8 and email clients. | Produce printable HTML tables with consistent layout by collapsing borders and enabling IE compatibility mode. | Build automated HTML email templates from Excel data that maintain visual fidelity across diverse client applications.
// AI Prompts: Write C# code using Aspose.Cells to export a worksheet to HTML with ExportSimilarBorderStyle, IsIECompatible, and IsBorderCollapsed enabled, then verify the CSS contains a fallback border rule. | Explain how ExportSimilarBorderStyle interacts with IsIECompatible and IsBorderCollapsed in HtmlSaveOptions. | Create a unit test that asserts the saved HTML file includes a thin border style for a cell originally styled with CellBorderType.Double.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to configure Aspose.Cells for .NET to save a workbook as HTML with ExportSimilarBorderStyle, IE compatibility, and collapsed borders, so unsupported border types (e.g., double) fall back to a compatible style in legacy browsers.
    public class ExportSimilarBorderStyleDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data
                sheet.Cells["A1"].PutValue("Original Border");
                sheet.Cells["B1"].PutValue("Fallback Border");

                // Apply a border style that may not be supported by older browsers (Double)
                Style doubleBorderStyle = workbook.CreateStyle();
                doubleBorderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Double;
                doubleBorderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Double;
                doubleBorderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Double;
                doubleBorderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Double;
                sheet.Cells["A1"].SetStyle(doubleBorderStyle);

                // Apply a supported border style for comparison (Thin)
                Style thinBorderStyle = workbook.CreateStyle();
                thinBorderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                thinBorderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                thinBorderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                thinBorderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                sheet.Cells["B1"].SetStyle(thinBorderStyle);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Export a similar (fallback) border style when the original style is not supported
                    ExportSimilarBorderStyle = true,
                    // Enable IE compatibility mode to simulate legacy browsers
                    IsIECompatible = true,
                    // Keep borders collapsed for a cleaner table layout
                    IsBorderCollapsed = true
                };

                // Save the workbook as HTML using the configured options
                string outputPath = "ExportSimilarBorderStyle.html";
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point required by the project
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportSimilarBorderStyleDemo.Run();
        }
    }
}
