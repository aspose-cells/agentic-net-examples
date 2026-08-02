// Title: Add a ListBox dropdown to an Excel worksheet and bind it to a chart data source using Aspose.Cells for .NET
// Description: This example creates a workbook, populates cells A1:B5 with categories and values, adds a column chart that reads its series from cell C1, inserts a ListBox shape linked to the range B2:B5, connects the ListBox to cell C1, updates the linked cell when a selection changes, and saves the file as DropdownChartDemo.xlsx.
// Keywords: Aspose.Cells | C# | Excel dropdown list | ListBox shape | bind dropdown to chart | linked cell | column chart | interactive Excel workbook | Aspose.Cells example | chart data source
// Common Searches: Aspose.Cells bind ListBox to chart series | C# add dropdown list to Excel worksheet | link ListBox selected value to chart data | Aspose.Cells update chart from dropdown | create interactive chart with dropdown in .NET
// Developer Intent: Create a ListBox dropdown on a worksheet, link its selected value to a single cell, and have a chart automatically reflect that value.
// Use Cases: Sales dashboard where selecting a month updates a column chart. | Reporting template that lets users choose a threshold, instantly refreshing the chart. | Training workbook demonstrating how a ListBox can drive chart data without formulas.
// AI Prompts: Write C# code with Aspose.Cells to add a ListBox dropdown, set its input range, link it to a cell, and bind that cell to a chart series. | Explain how to trigger chart refresh when the linked cell of a ListBox changes in Aspose.Cells for .NET. | Provide step‑by‑step instructions for building an interactive Excel file where a ListBox selection updates a column chart using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsDropdownChartDemo
{
    // This example creates a workbook, populates cells A1:B5 with categories and values, adds a column chart that reads its series from cell C1, inserts a ListBox shape linked to the range B2:B5, connects the ListBox to cell C1, updates the linked cell when a selection changes, and saves the file as DropdownChartDemo.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Populate sample data for the chart (Category / Value)
                // -------------------------------------------------
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

                // -------------------------------------------------
                // Add a column chart linked to the sample data
                // -------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the series to use a single cell (C1) that will be driven by the dropdown
                sheet.Cells["C1"].PutValue(10); // initial value
                chart.NSeries.Add("=Sheet1!$C$1", true);
                // Category data line removed because the Series class in the referenced Aspose.Cells version
                // does not expose a CategoryData property. The chart will use default categories.

                // -------------------------------------------------
                // Add a ListBox (dropdown) to the worksheet
                // -------------------------------------------------
                // Position: topRow=1, top=0, leftColumn=3, left=0, height=100, width=80
                ListBox listBox = sheet.Shapes.AddListBox(1, 0, 3, 0, 100, 80);

                // Use the same values as the chart (B2:B5) as the dropdown items
                listBox.SetInputRange("B2:B5", false, false);

                // Link the selected item to cell C1 (the chart's data source)
                listBox.LinkedCell = "C1";

                // Optionally set a default selected index (0 = first item)
                listBox.SelectedIndex = 0;

                // Update the linked cell with the selected value
                sheet.Shapes.UpdateSelectedValue();

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                string outputPath = "DropdownChartDemo.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
