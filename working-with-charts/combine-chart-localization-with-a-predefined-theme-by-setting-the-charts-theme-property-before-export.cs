// Title: Aspose.Cells C# – Apply a Custom Workbook Theme and Chinese Localization to a Chart, Then Export to PDF
// Description: This example demonstrates how to create a workbook, define a corporate color palette with Workbook.CustomTheme, configure chart globalization (titles, axis labels, display units) in Chinese using SettableChartGlobalizationSettings, add a column chart that inherits the theme, and generate a PDF of the chart while also saving the workbook.
// Keywords: Aspose.Cells | C# | custom workbook theme | chart localization | SettableChartGlobalizationSettings | Chinese chart titles | export chart to PDF | Workbook.CustomTheme | chart theme property | globalization settings
// Common Searches: Aspose.Cells apply custom theme to chart | C# set Chinese titles on Aspose.Cells chart | export Aspose.Cells chart as PDF with theme | SettableChartGlobalizationSettings example | Workbook.CustomTheme usage in .NET
// Developer Intent: Combine a predefined workbook theme with Chinese chart globalization and export the chart as PDF.
// Use Cases: Produce a sales‑report column chart that follows corporate colors and displays Chinese labels, then share it as a PDF. | Create a template workbook where multiple charts inherit the same corporate palette and localization for consistent regional branding. | Maintain an editable Excel file with the defined palette while extracting individual charts as PDFs for presentation.
// AI Prompts: Generate C# code using Aspose.Cells to define a custom color theme, set Chinese titles via SettableChartGlobalizationSettings, apply the theme to a chart, and export the chart to PDF. | Demonstrate how to configure Aspose.Cells chart globalization for axis titles, display units, and series names, then bind a predefined workbook theme before calling Chart.ToPdf. | Provide a complete example that creates a column chart from sample data, applies a custom theme and Chinese localization, and saves both the chart PDF and the workbook file.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartThemeLocalization
{
    // Custom globalization settings that use SettableChartGlobalizationSettings
    // This example demonstrates how to create a workbook, define a corporate color palette with Workbook.CustomTheme, configure chart globalization (titles, axis labels, display units) in Chinese using SettableChartGlobalizationSettings, add a column chart that inherits the theme, and generate a PDF of the chart while also saving the workbook.
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        public CustomGlobalizationSettings(SettableChartGlobalizationSettings chartSettings)
        {
            this.ChartSettings = chartSettings;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Apply a predefined custom theme to the workbook
            // -------------------------------------------------
            Color[] customColors = new Color[]
            {
                Color.FromArgb(255, 255, 255), // Background1
                Color.FromArgb(0, 0, 0),       // Text1
                Color.FromArgb(240, 240, 240), // Background2
                Color.FromArgb(80, 80, 80),    // Text2
                Color.FromArgb(0, 112, 192),   // Accent1
                Color.FromArgb(255, 192, 0),   // Accent2
                Color.FromArgb(112, 173, 71),  // Accent3
                Color.FromArgb(255, 0, 0),     // Accent4
                Color.FromArgb(0, 176, 80),    // Accent5
                Color.FromArgb(0, 176, 240),   // Accent6
                Color.FromArgb(0, 0, 255),     // Hyperlink
                Color.FromArgb(128, 0, 128)    // Followed Hyperlink
            };
            workbook.CustomTheme("PredefinedTheme", customColors);

            // -------------------------------------------------
            // 2. Configure chart localization (globalization)
            // -------------------------------------------------
            var chartGlobals = new SettableChartGlobalizationSettings();
            chartGlobals.SetChartTitleName("销售报告");          // "Sales Report" in Chinese
            chartGlobals.SetAxisTitleName("月份");              // "Month"
            chartGlobals.SetAxisUnitName(DisplayUnitType.Thousands, "千"); // "Thousand"
            chartGlobals.SetSeriesName("收入");                // "Revenue"

            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings(chartGlobals);

            // -------------------------------------------------
            // 3. Populate sample data for the chart
            // -------------------------------------------------
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Revenue");
            string[] months = { "Jan", "Feb", "Mar", "Apr" };
            int[] revenue = { 12000, 15000, 18000, 21000 };

            for (int i = 0; i < months.Length; i++)
            {
                sheet.Cells[i + 2, 0].PutValue(months[i]);   // Column A
                sheet.Cells[i + 2, 1].PutValue(revenue[i]); // Column B
            }

            // -------------------------------------------------
            // 4. Add a chart and apply the theme
            // -------------------------------------------------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set data range for the chart
            chart.SetChartDataRange("A1:B5", true);

            // Apply built‑in style (optional, demonstrates theme usage)
            chart.Style = 7; // any style between 1‑48

            // -------------------------------------------------
            // 5. Export the chart to PDF (chart inherits workbook theme)
            // -------------------------------------------------
            string pdfPath = "ChartWithThemeAndLocalization.pdf";
            chart.ToPdf(pdfPath);

            // Also save the workbook for reference
            workbook.Save("WorkbookWithThemeAndLocalization.xlsx");
        }
    }
}
