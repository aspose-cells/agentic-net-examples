// Title: Create identical column charts on every worksheet of an Excel file using Aspose.Cells for .NET
// Common Searches: how to add the same column chart to each sheet with Aspose.Cells C# | loop through all worksheets and insert a column chart using Aspose.Cells .NET | batch generate identical charts across multiple Excel worksheets in C# | apply a chart template to every worksheet programmatically with Aspose.Cells
// Tags: batch add column chart Aspose.Cells | iterate worksheets for chart insertion C# | chart series range binding Aspose.Cells | apply chart template across workbook .NET | programmatic Excel chart creation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing workbook or creates a new one with sample data, then iterates over each worksheet, adds a column chart positioned from cell (0,0) to (15,5), sets the chart title to "Sales Overview", binds the series to the range A1:B5 using the first column as series names, optionally assigns the series name from cell A1, and saves the modified workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if present; otherwise create a new one with sample data
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue("Month");
                ws.Cells["B1"].PutValue("Sales");
                ws.Cells["A2"].PutValue("Jan");
                ws.Cells["B2"].PutValue(100);
                ws.Cells["A3"].PutValue("Feb");
                ws.Cells["B3"].PutValue(150);
                ws.Cells["A4"].PutValue("Mar");
                ws.Cells["B4"].PutValue(200);
                ws.Cells["A5"].PutValue("Apr");
                ws.Cells["B5"].PutValue(250);
            }

            // Add a column chart to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int chartIndex = sheet.Charts.Add(ChartType.Column, 0, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                chart.Title.Text = "Sales Overview";

                // Use range A1:B5; first column contains series name
                chart.NSeries.Add("A1:B5", true);

                // Set series name from the first cell if available
                if (chart.NSeries.Count > 0)
                {
                    chart.NSeries[0].Name = sheet.Cells["A1"].StringValue;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
