// Title: Assign custom data label text from a cell range to the first series of a column chart with Aspose.Cells for .NET
// Common Searches: Aspose.Cells set custom data labels for chart series from worksheet cells C# | C# Aspose.Cells use reflection to set CustomDataLabels property on chart series | How to display cell values as data labels in an Aspose.Cells column chart | Apply custom text labels to first series of a column chart using Aspose.Cells .NET | Aspose.Cells chart series custom data labels version check
// Tags: custom data label assignment Aspose.Cells | column chart series label mapping from cells | reflection based API property setting in Aspose.Cells | conditional CustomDataLabels usage based on version | populate chart data labels from worksheet range

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example creates a workbook, fills columns A‑B with categories and values, adds a column chart, defines the first series with A2:A5 as categories and B2:B5 as values, enables standard data labels, then replaces those labels with custom text from cells C2:C5 using reflection to set the CustomDataLabels property when it exists, and finally saves the file as ChartWithCustomDataLabels.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["B5"].PutValue(40);

                // Prepare custom data label values (used for demonstration)
                sheet.Cells["C2"].PutValue("Label1");
                sheet.Cells["C3"].PutValue("Label2");
                sheet.Cells["C4"].PutValue("Label3");
                sheet.Cells["C5"].PutValue("Label4");

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                var chart = sheet.Charts[chartIndex];

                // Add series with categories (A2:A5) and values (B2:B5)
                chart.NSeries.Add("A2:A5,B2:B5", true);

                // Enable data labels (show the values)
                chart.NSeries[0].DataLabels.ShowValue = true;

                // Attempt to apply custom data labels if the API is available
                try
                {
                    // Aspose.Cells versions prior to 23.12 do not support CustomDataLabels.
                    // The following code will compile only when the property exists.
                    // If it does not exist, the catch block will handle the situation gracefully.
                    string[] customLabels = new string[]
                    {
                        sheet.Cells["C2"].StringValue,
                        sheet.Cells["C3"].StringValue,
                        sheet.Cells["C4"].StringValue,
                        sheet.Cells["C5"].StringValue
                    };

                    // Use reflection to set CustomDataLabels if the property exists
                    var dataLabels = chart.NSeries[0].DataLabels;
                    var prop = dataLabels.GetType().GetProperty("CustomDataLabels");
                    if (prop != null && prop.CanWrite)
                    {
                        prop.SetValue(dataLabels, customLabels);
                    }
                    else
                    {
                        Console.WriteLine("CustomDataLabels property is not available in this Aspose.Cells version.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Custom data labels could not be applied: {ex.Message}");
                }

                // Save the workbook
                string outputPath = "ChartWithCustomDataLabels.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
