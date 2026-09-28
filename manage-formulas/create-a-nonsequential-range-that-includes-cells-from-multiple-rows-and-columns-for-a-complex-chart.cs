// Title: Generate a line chart with non‑adjacent data ranges (A2:A5, C2:C5, E2:E5) using Aspose.Cells for .NET (C#)
// AI Prompts: Create a C# program that uses Aspose.Cells to build a line chart where the X‑axis values come from A2:A5 and two Y‑series are taken from C2:C5 and E2:E5, and set the series names from the header cells. | Show how to add multiple series to an Aspose.Cells chart with separate column ranges and assign XValues and Name properties programmatically. | Demonstrate saving the workbook with the constructed chart to an .xlsx file after populating sample data in columns A, C, and E.
// Common Searches: how to plot a line chart from non‑contiguous columns using Aspose.Cells C# | Aspose.Cells set XValues for chart series from a different column | use header cell as series name in Aspose.Cells chart example | create chart with multiple series from separate columns Aspose.Cells .NET | non‑sequential range chart Aspose.Cells line chart tutorial
// Tags: Aspose.Cells create line chart multiple column ranges | Aspose.Cells configure series XValues | Aspose.Cells assign series name from cell reference | Aspose.Cells populate worksheet data across columns | Aspose.Cells save workbook with chart to XLSX

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsNonSequentialRangeChart
{
    // The sample creates a workbook, fills columns A, C, and E with category and series data, adds a line chart, assigns X‑values from A2:A5 and Y‑values from C2:C5 and E2:E5, uses the header cells as series names, sets a chart title, and saves the file as NonSequentialRangeChart.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data in multiple columns (A, C, E) and rows (1-5)

                // Column A - Categories
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");

                // Column C - Series 1 values
                sheet.Cells["C1"].PutValue("Series1");
                sheet.Cells["C2"].PutValue(10);
                sheet.Cells["C3"].PutValue(20);
                sheet.Cells["C4"].PutValue(30);
                sheet.Cells["C5"].PutValue(40);

                // Column E - Series 2 values
                sheet.Cells["E1"].PutValue("Series2");
                sheet.Cells["E2"].PutValue(15);
                sheet.Cells["E3"].PutValue(25);
                sheet.Cells["E4"].PutValue(35);
                sheet.Cells["E5"].PutValue(45);

                // Add a line chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Line, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Add first series (Series1) with Y values from C2:C5
                chart.NSeries.Add("C2:C5", true);
                // Set X values (categories) for the first series
                chart.NSeries[0].XValues = "A2:A5";

                // Add second series (Series2) with Y values from E2:E5
                chart.NSeries.Add("E2:E5", true);
                // Set X values (categories) for the second series
                chart.NSeries[1].XValues = "A2:A5";

                // Use header cells as series names
                chart.NSeries[0].Name = "='Sheet1'!$C$1";
                chart.NSeries[1].Name = "='Sheet1'!$E$1";

                // Set chart title
                chart.Title.Text = "Complex Line Chart with Non‑Sequential Ranges";

                // Save the workbook
                workbook.Save("NonSequentialRangeChart.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
