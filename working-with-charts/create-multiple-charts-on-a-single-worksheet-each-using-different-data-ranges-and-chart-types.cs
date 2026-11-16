// Title: Add a column chart and a pie chart with separate data ranges to the same worksheet using Aspose.Cells for .NET
// Common Searches: how to add a column chart and a pie chart to the same sheet using Aspose.Cells C# | Aspose.Cells C# set different data ranges for multiple charts in one worksheet | position multiple charts on a worksheet with Aspose.Cells .NET | assign titles to individual charts in an Aspose.Cells workbook | generate Excel file with both column and pie charts using Aspose.Cells
// Tags: Aspose.Cells add multiple charts to worksheet | column chart custom data range Aspose.Cells | pie chart positioning Aspose.Cells C# | set chart titles Aspose.Cells workbook | save workbook with multiple charts Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsMultipleCharts
{
    // The example creates a new workbook, fills two separate data tables, adds a column chart and a pie chart on the same worksheet with distinct data ranges and layout positions, sets titles for each chart, and saves the file as MultipleCharts.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data for the first chart (Column Chart)
                // Category (X) values in column A, Series values in column B
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["B5"].PutValue(40);

                // Populate data for the second chart (Pie Chart)
                // Category (X) values in column D, Series values in column E
                sheet.Cells["D1"].PutValue("Category");
                sheet.Cells["E1"].PutValue("Series2");
                sheet.Cells["D2"].PutValue("Apple");
                sheet.Cells["D3"].PutValue("Banana");
                sheet.Cells["D4"].PutValue("Cherry");
                sheet.Cells["D5"].PutValue("Date");
                sheet.Cells["E2"].PutValue(25);
                sheet.Cells["E3"].PutValue(35);
                sheet.Cells["E4"].PutValue(15);
                sheet.Cells["E5"].PutValue(25);

                // ------------------------------
                // Create the first chart (Column)
                // ------------------------------
                // Position: rows 7-20, columns 0-7 (A8:H21)
                int chart1Index = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 7);
                Chart chart1 = sheet.Charts[chart1Index];

                // Set the data series for the column chart
                chart1.NSeries.CategoryData = "A2:A5";
                chart1.NSeries.Add("B2:B5", true);

                // Optional: set chart title
                chart1.Title.Text = "Monthly Sales";

                // ------------------------------
                // Create the second chart (Pie)
                // ------------------------------
                // Position: rows 7-20, columns 9-16 (J8:Q21)
                int chart2Index = sheet.Charts.Add(ChartType.Pie, 7, 9, 20, 16);
                Chart chart2 = sheet.Charts[chart2Index];

                // Set the data series for the pie chart
                chart2.NSeries.CategoryData = "D2:D5";
                chart2.NSeries.Add("E2:E5", true);

                // Optional: set chart title
                chart2.Title.Text = "Fruit Distribution";

                // Determine output file path
                string outputPath = "MultipleCharts.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
