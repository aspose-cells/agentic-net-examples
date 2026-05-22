using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AxisLabelTranslationDemo
{
    class Program
    {
        // Placeholder for an external translation service.
        // In a real scenario this would call a web API or other service.
        static string TranslateLabel(string original)
        {
            // Example: prepend "TR-" to simulate translation.
            return "TR-" + original;
        }

        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Populate sample data ----------
            // Category (axis labels) in column A
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Cherry");

            // Values in column B
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["B4"].PutValue(150);

            // ---------- Add a column chart ----------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories (axis labels)

            // ---------- Calculate chart to generate axis texts ----------
            chart.Calculate();

            // ---------- Retrieve current axis labels ----------
            string[] originalLabels = chart.CategoryAxis.GetAxisTexts();

            // ---------- Translate each label ----------
            List<string> translatedLabels = new List<string>();
            foreach (string label in originalLabels)
            {
                translatedLabels.Add(TranslateLabel(label));
            }

            // ---------- Update the source cells with translated labels ----------
            // Assuming the category data range starts at A2 and has the same count as originalLabels.
            for (int i = 0; i < translatedLabels.Count; i++)
            {
                // Row index = 2 + i (since A2 is the first data cell)
                int row = 2 + i;
                sheet.Cells[row, 0].PutValue(translatedLabels[i]); // Column 0 = A
            }

            // ---------- Re‑calculate the chart so it reflects the new labels ----------
            chart.Calculate();

            // Optional: Update axis title to indicate translation
            chart.CategoryAxis.Title.Text = "Translated Categories";

            // ---------- Save the workbook ----------
            workbook.Save("AxisLabelTranslationDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}