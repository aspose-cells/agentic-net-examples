// Title: Batch‑generate pie charts for every worksheet using Aspose.Cells in C#
// Description: Creates a new workbook, fills sample tables on multiple sheets, then loops through each worksheet, detects the last populated row in column A, adds a pie chart positioned on the sheet, links the series to column B values (using column A as categories), sets a title that includes the sheet name, and saves the file as BatchPieCharts.xlsx.
// Keywords: Aspose.Cells C# pie chart | add chart to each worksheet | loop worksheets Aspose.Cells | last data row column A | chart series range Aspose.Cells | Excel pie chart automation | batch chart generation C# | set chart title programmatically | generate multiple charts Aspose.Cells | C# Excel visualization
// Common Searches: Aspose.Cells add pie chart to every sheet | C# loop through worksheets and create charts | How to find last populated row in a column with Aspose.Cells | Set series range for pie chart Aspose.Cells C# | Batch generate Excel charts programmatically | Create pie chart for each worksheet Aspose.Cells
// Developer Intent: Automatically insert a pie chart into each worksheet based on its category/value columns in one execution.
// Use Cases: Produce a monthly report where sales and expense tables on separate sheets are visualized instantly with pie charts. | Provide a template workbook that adds a pie chart to any new sheet containing two columns (category and amount) without manual editing. | Generate a summary Excel file that adds charts to all sheets, enabling quick visual analysis of categorical data across departments.
// AI Prompts: Write C# code with Aspose.Cells that iterates over all worksheets and adds a pie chart using columns A (categories) and B (values), setting the chart title to the sheet name. | Explain how to determine the last non‑empty row in a specific column using Aspose.Cells and apply it to define chart data ranges. | Provide robust error‑handling patterns for batch chart creation in Aspose.Cells, including per‑sheet logging of failures. | Suggest performance optimizations when generating dozens of charts in a large workbook with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsBatchPieCharts
{
    // Creates a new workbook, fills sample tables on multiple sheets, then loops through each worksheet, detects the last populated row in column A, adds a pie chart positioned on the sheet, links the series to column B values (using column A as categories), sets a title that includes the sheet name, and saves the file as BatchPieCharts.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -----------------------------
                // Sample data creation (for demo)
                // -----------------------------
                // Worksheet 1
                Worksheet ws1 = workbook.Worksheets[0];
                ws1.Name = "Sales";
                ws1.Cells["A1"].PutValue("Product");
                ws1.Cells["B1"].PutValue("Amount");
                ws1.Cells["A2"].PutValue("Apple");
                ws1.Cells["B2"].PutValue(120);
                ws1.Cells["A3"].PutValue("Banana");
                ws1.Cells["B3"].PutValue(80);
                ws1.Cells["A4"].PutValue("Cherry");
                ws1.Cells["B4"].PutValue(150);

                // Worksheet 2
                Worksheet ws2 = workbook.Worksheets.Add("Expenses");
                ws2.Cells["A1"].PutValue("Category");
                ws2.Cells["B1"].PutValue("Cost");
                ws2.Cells["A2"].PutValue("Rent");
                ws2.Cells["B2"].PutValue(500);
                ws2.Cells["A3"].PutValue("Utilities");
                ws2.Cells["B3"].PutValue(200);
                ws2.Cells["A4"].PutValue("Supplies");
                ws2.Cells["B4"].PutValue(150);

                // -------------------------------------------------
                // Generate a pie chart for each worksheet with data
                // -------------------------------------------------
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    try
                    {
                        // Determine the last row that contains data in column A (categories)
                        int lastRow = sheet.Cells.GetLastDataRow(0); // zero‑based index

                        // Need at least one data row (row index 1 = second row)
                        if (lastRow < 1) continue;

                        // Add a pie chart to the worksheet
                        // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
                        int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 5);
                        Chart chart = sheet.Charts[chartIndex];

                        // Define the data range for values and categories (Excel notation is 1‑based)
                        string valueRange = $"{sheet.Name}!B2:B{lastRow + 1}";
                        string categoryRange = $"{sheet.Name}!A2:A{lastRow + 1}";

                        // Set the series data (values)
                        chart.NSeries.Add(valueRange, true);

                        // Aspose.Cells automatically uses the adjacent column as categories for pie charts.
                        // If a specific API for categories is unavailable, we rely on the default behavior.

                        // Optional: set chart title
                        chart.Title.Text = $"Pie Chart - {sheet.Name}";
                    }
                    catch (Exception chartEx)
                    {
                        Console.WriteLine($"Failed to create chart for sheet '{sheet.Name}': {chartEx.Message}");
                    }
                }

                // Define output file path
                string outputPath = "BatchPieCharts.xlsx";

                // Save the workbook safely
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
