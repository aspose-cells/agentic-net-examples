// Title: Export Excel Charts to Separate PDFs with Custom Globalization using Aspose.Cells (C#)
// Description: Load an Excel workbook, apply a CustomGlobalizationSettings object that overrides chart titles, legends and axis units, then iterate every worksheet and export each chart to its own PDF file. Includes file‑existence check and error handling.
// Keywords: Aspose.Cells | C# | export chart to PDF | chart globalization | CustomGlobalizationSettings | ChartGlobalizationSettings | Excel chart PDF | separate PDF per chart | localization | globalization | Aspose.Cells API | Chart.ToPdf
// Common Searches: How to export each Excel chart to a separate PDF with Aspose.Cells | Apply custom chart globalization when converting charts to PDF in C# | Aspose.Cells export chart with custom titles and axis labels | Batch convert workbook charts to PDFs using Aspose.Cells | C# code for chart ToPdf with custom GlobalizationSettings
// Developer Intent: Generate individual PDF files for all workbook charts while applying custom globalization strings.
// Use Cases: Produce localized PDF reports for financial dashboards where each chart needs language‑specific titles and legends. | Automate creation of separate PDF assets for marketing presentations, customizing chart captions per region. | Integrate chart‑to‑PDF conversion into a CI pipeline that respects custom globalization for multi‑language releases.
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, sets a CustomGlobalizationSettings object, and exports every chart to a distinct PDF file. | Show how to subclass ChartGlobalizationSettings to override series names, titles, legends, and axis units, then use it during chart PDF export. | Explain how to safely check for a missing workbook file, log each exported PDF name, and handle exceptions while converting charts with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Load an Excel workbook, apply a CustomGlobalizationSettings object that overrides chart titles, legends and axis units, then iterate every worksheet and export each chart to its own PDF file. Includes file‑existence check and error handling.
class ExportChartsToPdf
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Apply custom globalization settings with custom chart settings
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings
            {
                ChartSettings = new CustomChartGlobalizationSettings()
            };

            // Export each chart in each worksheet to a separate PDF file
            for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
            {
                Worksheet sheet = workbook.Worksheets[sheetIndex];
                for (int chartIndex = 0; chartIndex < sheet.Charts.Count; chartIndex++)
                {
                    Chart chart = sheet.Charts[chartIndex];
                    string pdfFileName = $"Chart_Sheet{sheetIndex}_Chart{chartIndex}.pdf";
                    chart.ToPdf(pdfFileName);
                    Console.WriteLine($"Exported chart to '{pdfFileName}'.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Custom globalization settings (no additional overrides needed for workbook level)
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
    }

    // Custom chart globalization settings to demonstrate overriding chart text
    public class CustomChartGlobalizationSettings : ChartGlobalizationSettings
    {
        public override string GetSeriesName()
        {
            return "Custom Series";
        }

        public override string GetChartTitleName()
        {
            return "Custom Chart Title";
        }

        public override string GetLegendIncreaseName()
        {
            return "Custom Increase";
        }

        public override string GetLegendDecreaseName()
        {
            return "Custom Decrease";
        }

        public override string GetOtherName()
        {
            return "Custom Other";
        }

        public override string GetAxisTitleName()
        {
            return "Custom Axis Title";
        }

        public override string GetAxisUnitName(DisplayUnitType type)
        {
            // Example: customize unit names
            return type switch
            {
                DisplayUnitType.Hundreds => "Hundreds_Custom",
                DisplayUnitType.Thousands => "Thousands_Custom",
                DisplayUnitType.TenThousands => "TenThousands_Custom",
                _ => base.GetAxisUnitName(type),
            };
        }

        public override string GetLegendTotalName()
        {
            return "Custom Total";
        }
    }
}
