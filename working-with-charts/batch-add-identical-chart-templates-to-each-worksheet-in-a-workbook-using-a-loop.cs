using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (contains a default worksheet)
            Workbook workbook = new Workbook();

            // Remove the default worksheet to avoid name conflicts
            if (workbook.Worksheets.Count > 0)
                workbook.Worksheets.RemoveAt(0);

            // Add worksheets with unique names
            workbook.Worksheets.Add("Sheet1");
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Verify that the chart template file exists
            const string templatePath = "ChartTemplate.crtx";
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file \"{templatePath}\" not found.");
                return;
            }

            // Load the chart template (.crtx) into a byte array
            byte[] templateData = File.ReadAllBytes(templatePath);

            // Loop through each worksheet and add the same chart template
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Populate sample data that the chart will use (range A1:B5)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                for (int i = 1; i <= 4; i++)
                {
                    sheet.Cells[$"A{i + 1}"].PutValue($"Item {i}");
                    sheet.Cells[$"B{i + 1}"].PutValue(i * 10);
                }

                // Add a chart using the template.
                // Parameters: template bytes, data range, isVertical, topRow, leftColumn, bottomRow, rightColumn
                int chartIndex = sheet.Charts.Add(templateData, "A1:B5", true, 5, 0, 20, 7);
                Chart chart = sheet.Charts[chartIndex];

                // Optional: set a title to identify the sheet
                chart.Title.Text = $"Chart on {sheet.Name}";
            }

            // Save the workbook with the added charts
            workbook.Save("WorkbookWithCharts.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}