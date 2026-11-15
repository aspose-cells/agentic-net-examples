// Title: Check default series name and empty axis titles in an Aspose.Cells chart when no localization is applied (C#)
// Common Searches: how to confirm default series name is 'Series 1' in Aspose.Cells chart C# | verify that chart axis titles are blank when localization is not set using Aspose.Cells | Aspose.Cells chart default English text without localization example C# | C# code to test chart NSeries default name in Aspose.Cells workbook | determine if chart axis titles are empty in Aspose.Cells when no locale is assigned
// Tags: Aspose.Cells chart default series name C# | empty axis titles Aspose.Cells chart | chart localization default behavior Aspose.Cells | verify chart NSeries name Aspose.Cells | C# Aspose.Cells chart without locale

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample creates a workbook, adds sample data and a column chart, then asserts that the first series name defaults to "Series 1" and that both category and value axis titles are empty, confirming that no localization text is applied; finally the workbook is saved to a memory stream.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Add sample data to the first worksheet
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["A3"].PutValue("B");
            ws.Cells["A4"].PutValue("C");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["B3"].PutValue(20);
            ws.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIdx];

            // Set series data
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Verify default series name (should be "Series 1")
            string seriesName = chart.NSeries[0].Name;
            if (seriesName != "Series 1")
                throw new Exception($"Expected default series name 'Series 1', but got '{seriesName}'.");

            // Verify that axis titles are empty (no localization applied)
            string categoryAxisTitle = chart.CategoryAxis.Title.Text;
            string valueAxisTitle = chart.ValueAxis.Title.Text;
            if (!string.IsNullOrEmpty(categoryAxisTitle) || !string.IsNullOrEmpty(valueAxisTitle))
                throw new Exception("Axis titles should be empty when no localization is assigned.");

            // Save workbook to a memory stream (demonstrates lifecycle usage)
            using (MemoryStream ms = new MemoryStream())
            {
                wb.Save(ms, SaveFormat.Xlsx);
            }

            Console.WriteLine("Default English text verification passed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
