// Title: Add a column chart and freeze the first five rows in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create column chart and freeze top rows in worksheet | C# code to add chart and apply FreezePanes to first rows using Aspose.Cells | How to set chart series range A2:B5 and freeze rows with Aspose.Cells .NET | Example of freezing panes after inserting a chart in an Excel file with Aspose.Cells | Saving Excel file with chart and frozen rows using Aspose.Cells for .NET
// Tags: column chart creation Aspose.Cells C# | freeze panes first rows Aspose.Cells | chart series range A2:B5 Aspose.Cells | chart insertion with frozen rows Aspose.Cells | Excel workbook save with chart Aspose.Cells | set chart title Aspose.Cells C#

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample program creates a new workbook, fills cells A1:B5 with sample data, adds a column chart positioned from rows 7 to 20, sets the chart title, defines the series using the range A2:B5, freezes the first five rows to keep the data visible, and saves the file as ChartWithFrozenRows.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["A5"].PutValue("Apr");
            sheet.Cells["B5"].PutValue(200);

            // Add a column chart that uses the data range (including categories)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Monthly Sales";

            // Define the series with both categories (A2:A5) and values (B2:B5)
            chart.NSeries.Add("A2:B5", true);

            // Freeze the top 5 rows (rows 0‑4). Use the overload that specifies frozen rows.
            sheet.FreezePanes(0, 0, 5, 0);

            // Prepare output path
            string outputPath = "ChartWithFrozenRows.xlsx";
            string fullPath = Path.GetFullPath(outputPath);
            string directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {fullPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
