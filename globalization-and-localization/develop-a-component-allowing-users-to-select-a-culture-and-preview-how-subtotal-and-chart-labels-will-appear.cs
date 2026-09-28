// Title: Create a culture‑aware Excel preview with subtotal row and localized chart labels using Aspose.Cells for .NET
// AI Prompts: Write a C# method that builds an XLSX workbook, populates it with sample category and amount data, applies the number format of a supplied CultureInfo to the amount column, adds a subtotal row, inserts a column chart with data labels, and returns the workbook as a byte array. | Enhance the method to set the chart title to include the culture name and configure the chart data labels to use the same culture‑specific number format.
// Common Searches: asp.net apply CultureInfo number format to Excel column using Aspose.Cells | c# generate Excel file with subtotal row and localized chart labels | how to preview chart labels for different cultures in Aspose.Cells | create culture specific Excel preview with column chart in .NET | asp.net Aspose.Cells format numbers per locale in generated workbook
// Tags: culture-specific number formatting Aspose.Cells | add subtotal row to workbook Aspose.Cells | insert column chart with data labels Aspose.Cells | generate Excel preview as byte array C# | localize chart title using CultureInfo Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The CulturePreviewComponent builds an XLSX workbook, fills it with sample category and amount data, applies a custom number format based on a provided CultureInfo, adds a subtotal row, creates a column chart with data labels, sets the chart title to include the culture name, and returns the workbook saved as a byte array.
public class CulturePreviewComponent
{
    /// <param name="cultureName">The name of the culture (e.g., "en-US", "fr-FR").</param>
    /// <returns>A byte array containing the generated Excel file.</returns>
    public byte[] GeneratePreview(string cultureName)
    {
        try
        {
            // Validate culture
            CultureInfo culture;
            try
            {
                culture = new CultureInfo(cultureName);
            }
            catch (CultureNotFoundException)
            {
                throw new ArgumentException($"Culture '{cultureName}' is not valid.", nameof(cultureName));
            }

            // Create a new workbook and get the first worksheet cells
            var workbook = new Workbook();
            var cells = workbook.Worksheets[0].Cells;

            // Header
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");

            // Sample data
            string[] categories = { "Food", "Food", "Transport", "Transport", "Utilities", "Utilities" };
            double[] amounts = { 120.5, 85.75, 60.0, 45.25, 150.0, 130.5 };

            for (int i = 0; i < categories.Length; i++)
            {
                cells[i + 1, 0].PutValue(categories[i]); // Column A (row index starts at 0)
                cells[i + 1, 1].PutValue(amounts[i]);   // Column B
            }

            // Add a simple total row (optional, replaces Subtotal feature)
            int totalRowIndex = categories.Length + 1;
            cells[totalRowIndex, 0].PutValue("Total");
            cells[totalRowIndex, 1].Formula = $"=SUM(B2:B{categories.Length + 1})";

            // Format the Amount column according to the selected culture
            Style amountStyle = workbook.CreateStyle();
            amountStyle.Custom = "#,##0.00";
            StyleFlag flag = new StyleFlag { NumberFormat = true };
            cells.Columns[1].ApplyStyle(amountStyle, flag);

            // Add a column chart
            int chartIndex = workbook.Worksheets[0].Charts.Add(ChartType.Column, 8, 0, 20, 10);
            Chart chart = workbook.Worksheets[0].Charts[chartIndex];

            // Set data source (values)
            chart.NSeries.Add("=Sheet1!$B$2:$B$7", true);

            // Enable data labels if supported
            try
            {
                chart.NSeries[0].DataLabels.ShowValue = true;
                chart.NSeries[0].DataLabels.NumberFormat = "#,##0.00";
            }
            catch
            {
                // Ignore if not supported
            }

            // Chart title reflecting the culture
            chart.Title.Text = $"Amount by Category ({culture.Name})";

            // Save to memory stream
            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsx);
                return ms.ToArray();
            }
        }
        catch (Exception ex)
        {
            // Wrap any exception for clearer diagnostics
            throw new ApplicationException("Failed to generate culture preview.", ex);
        }
    }
}

// Simple entry point for testing the component
public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            var component = new CulturePreviewComponent();
            byte[] fileBytes = component.GeneratePreview("en-US");

            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CulturePreview.xlsx");

            // Ensure the directory exists before writing
            string? dir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            File.WriteAllBytes(outputPath, fileBytes);
            Console.WriteLine($"Excel file generated successfully at: {outputPath}");
        }
        catch (Exception e)
        {
            Console.WriteLine($"Error: {e.Message}");
        }
    }
}
