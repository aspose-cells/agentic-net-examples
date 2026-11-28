// Title: How to enable automatic Y‑axis display units for a column chart using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set Y axis to auto display units in column chart | C# chart value axis automatic scaling with Aspose.Cells | Enable automatic display unit label for Y axis in Aspose.Cells chart | Aspose.Cells chart Y axis auto units example C# | Let Aspose.Cells choose appropriate Y axis unit for large numbers in a chart
// Tags: auto y axis scaling Aspose.Cells | chart y axis auto units C# | column chart y axis auto scaling Aspose.Cells | clear custom display unit Aspose.Cells chart | force chart calculation Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace EnableAutomaticYAxisUnits
{
    // The example creates a workbook, adds sample data, and builds a column chart. It accesses the chart's value (Y) axis and explicitly enables automatic minimum, maximum, major, and minor values while clearing any custom display unit. After forcing chart calculation, Aspose.Cells automatically determines the appropriate display unit (e.g., thousands) and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data that will require a display unit (e.g., thousands)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");
                sheet.Cells["B2"].PutValue(15000);
                sheet.Cells["B3"].PutValue(23000);
                sheet.Cells["B4"].PutValue(18000);
                sheet.Cells["B5"].PutValue(27000);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Access the value (Y) axis
                Axis yAxis = chart.ValueAxis;

                // Enable automatic handling of axis limits and units (default values, but set explicitly)
                yAxis.IsAutomaticMinValue = true;
                yAxis.IsAutomaticMaxValue = true;
                yAxis.IsAutomaticMajorUnit = true;
                yAxis.IsAutomaticMinorUnit = true;

                // Ensure no custom display unit is forced
                yAxis.DisplayUnit = DisplayUnitType.None;
                yAxis.IsDisplayUnitLabelShown = false;

                // Force chart calculation so that axis texts are generated
                chart.Calculate();

                // After calculation, Aspose.Cells may automatically decide a display unit.
                // If it does, the DisplayUnitLabel will contain the unit text.
                Console.WriteLine("Automatic Y‑Axis Display Unit Label: " +
                                  (yAxis.IsDisplayUnitLabelShown ? yAxis.DisplayUnitLabel.Text : "None (auto scaling)"));

                // Save the workbook
                string outFile = "AutomaticYAxisUnits.xlsx";
                workbook.Save(outFile, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
