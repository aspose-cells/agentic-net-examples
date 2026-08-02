// Title: C# – Update Accent5 Theme Color and Refresh All Charts with Aspose.Cells for .NET
// Description: Demonstrates how to set the workbook's Accent5 theme color to a user‑selected value using Aspose.Cells, trigger an internal refresh of every chart so they adopt the new theme, and save the updated file.
// Keywords: Aspose.Cells | SetThemeColor | Accent5 | theme color | refresh charts | C# | .NET | Excel workbook | chart update | user selected palette | Workbook.Save | Chart.Title | theme refresh
// Common Searches: change Accent5 color Aspose.Cells C# | refresh Excel charts after theme change .NET | set workbook theme color programmatically | Aspose.Cells update theme and charts | apply user palette to Excel workbook using Aspose
// Developer Intent: Set the Accent5 theme color to a custom RGB value and ensure every existing chart reflects the change without recreating the charts.
// Use Cases: Apply a corporate orange palette by updating Accent5 and automatically recoloring chart series across all worksheets. | Load an existing Excel file, modify its Accent5 theme color, trigger chart refresh, and save the file for distribution. | Batch‑process multiple workbooks to enforce a consistent theme and refresh all dependent visualizations.
// AI Prompts: Generate C# code with Aspose.Cells that changes the Accent5 theme color based on an RGB input and forces all charts to refresh. | Show how to update multiple theme accents (Accent1‑Accent6) in a workbook and make charts pick up the new colors. | Explain the mechanism for programmatically refreshing charts after a theme color change in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to set the workbook's Accent5 theme color to a user‑selected value using Aspose.Cells, trigger an internal refresh of every chart so they adopt the new theme, and save the updated file.
class UpdateAccent5Theme
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // new Workbook("input.xlsx");

        // Example user‑selected color for Accent5
        Color userSelectedColor = Color.FromArgb(255, 128, 0); // orange

        // Update the theme's Accent5 color
        workbook.SetThemeColor(ThemeColorType.Accent5, userSelectedColor);

        // Iterate through all worksheets and charts to ensure they pick up the new theme color
        foreach (Worksheet ws in workbook.Worksheets)
        {
            foreach (Chart chart in ws.Charts)
            {
                // Accessing a property forces internal refresh (no explicit refresh method)
                var _ = chart.Title.Text;
            }
        }

        // Save the modified workbook
        workbook.Save("UpdatedAccent5Theme.xlsx", SaveFormat.Xlsx);
    }
}
