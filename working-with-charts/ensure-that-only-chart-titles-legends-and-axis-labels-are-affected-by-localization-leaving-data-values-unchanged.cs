// Title: Translate Excel chart titles, legend entries, and axis labels with Aspose.Cells for .NET while preserving data values
// AI Prompts: Generate C# code that uses Aspose.Cells to iterate over every chart in a workbook and replace the text of the chart title, legend series names, and axis titles using a supplied dictionary, ensuring the underlying data series remain unchanged. | Demonstrate how to load an existing Excel file, apply a localization map to chart text elements (title, legend, axis labels) with Aspose.Cells, and save the workbook without affecting any cell values.
// Common Searches: Aspose.Cells C# change only chart title text in existing Excel file | How to localize legend series names in Excel charts using .NET without modifying data | Replace axis labels in Excel chart with dictionary lookup Aspose.Cells | Preserve chart data while translating chart labels in C# | Excel chart localization example Aspose.Cells .NET
// Tags: localize chart titles Aspose.Cells | replace legend series names .NET | axis label translation dictionary | preserve chart data values Aspose.Cells | Excel chart text localization C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads a workbook (or creates a new one), defines a dictionary mapping original strings to localized equivalents, walks through each worksheet's charts, and updates the chart title, each series name (legend), and the category and value axis titles using the dictionary. Cell data and series values are left untouched, and the modified workbook is saved.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load workbook if the file exists; otherwise create a new empty workbook.
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Simple localization map: original text -> localized text
            var localizationMap = new Dictionary<string, string>
            {
                { "Sales", "Ventas" },
                { "Revenue", "Ingresos" },
                { "Month", "Mes" },
                // Add more entries as needed
            };

            // Iterate through all worksheets and their charts
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    // ----- Localize Chart Title -----
                    if (chart.Title != null && !string.IsNullOrEmpty(chart.Title.Text))
                    {
                        chart.Title.Text = Localize(chart.Title.Text, localizationMap);
                    }

                    // ----- Localize Legend (Series Names) -----
                    foreach (Series series in chart.NSeries)
                    {
                        if (!string.IsNullOrEmpty(series.Name))
                        {
                            series.Name = Localize(series.Name, localizationMap);
                        }
                    }

                    // ----- Localize Axis Titles -----
                    // Category (X) Axis Title
                    if (chart.CategoryAxis != null && chart.CategoryAxis.Title != null && !string.IsNullOrEmpty(chart.CategoryAxis.Title.Text))
                    {
                        chart.CategoryAxis.Title.Text = Localize(chart.CategoryAxis.Title.Text, localizationMap);
                    }

                    // Value (Y) Axis Title
                    if (chart.ValueAxis != null && chart.ValueAxis.Title != null && !string.IsNullOrEmpty(chart.ValueAxis.Title.Text))
                    {
                        chart.ValueAxis.Title.Text = Localize(chart.ValueAxis.Title.Text, localizationMap);
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to replace text based on the localization map
    static string Localize(string original, Dictionary<string, string> map)
    {
        return map.TryGetValue(original, out string localized) ? localized : original;
    }
}
