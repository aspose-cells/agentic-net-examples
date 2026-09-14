// Title: Convert an Excel workbook to HTML with Aspose.Cells .NET while disabling tooltips and exporting only the active sheet
// AI Prompts: Generate HTML from a .xlsx file using Aspose.Cells in C# with AddTooltipText set to false. | Save only the active worksheet as HTML by configuring HtmlSaveOptions and turning off tooltip generation. | Speed up Excel‑to‑HTML conversion in C# by disabling tooltip text in Aspose.Cells HtmlSaveOptions.
// Common Searches: Aspose.Cells C# disable AddTooltipText when saving workbook as HTML | How to export only the active worksheet to HTML using Aspose.Cells .NET | Improve performance of Excel to HTML conversion by turning off tooltip text in Aspose.Cells | HtmlSaveOptions AddTooltipText false example C# | Export active sheet to HTML with Aspose.Cells without tooltips
// Tags: Aspose.Cells HtmlSaveOptions AddTooltipText false | export active worksheet only HTML Aspose.Cells | optimize Excel to HTML conversion performance .NET | disable tooltip generation Aspose.Cells HTML export | C# convert XLSX to HTML without tooltips

using Aspose.Cells;
using Aspose.Cells.Rendering;

// The C# program loads an Excel file, configures HtmlSaveOptions to disable tooltip text (AddTooltipText = false) and to export only the active worksheet, then saves the workbook as an HTML file using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load the source workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Disable tooltip text to improve rendering speed
            AddTooltipText = false,

            // Optional: export only the active worksheet to reduce processing time
            ExportActiveWorksheetOnly = true
        };

        // Save the workbook as an HTML file with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
