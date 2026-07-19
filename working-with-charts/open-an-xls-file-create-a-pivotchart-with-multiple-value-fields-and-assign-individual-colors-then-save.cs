// Title: C# – Create a PivotChart with Multiple Value Fields and Custom Series Colors from an XLS Workbook using Aspose.Cells
// Description: Loads an XLS file (or generates sample data), builds a PivotTable on A1:C10, adds Quantity and Revenue as data fields, creates a linked column PivotChart, assigns blue to the Quantity series and green to the Revenue series, sets a chart title, and saves the result as an XLSX workbook.
// Keywords: Aspose.Cells | C# | PivotChart | PivotTable | custom series colors | multiple value fields | load XLS | save XLSX | column chart | chart formatting | Excel automation
// Common Searches: Aspose.Cells set series color in PivotChart C# | Create PivotTable and linked PivotChart from XLS using Aspose.Cells | How to assign different colors to PivotChart series in .NET | C# example for PivotChart with multiple data fields Aspose.Cells | Load XLS and add colored PivotChart programmatically
// Developer Intent: Programmatically generate a PivotChart linked to a PivotTable, apply individual colors to each data series, and export the workbook.
// Use Cases: Generate a sales dashboard where quantity and revenue are visually distinguished by custom colors. | Automate monthly reporting by creating PivotTables from raw XLS data and adding brand‑compliant colored charts. | Export analytical Excel files that include colored PivotCharts for clearer metric comparison.
// AI Prompts: Show me C# code to create a PivotChart with separate colors for each series using Aspose.Cells. | How can I assign custom colors to PivotChart series after linking it to a PivotTable in Aspose.Cells? | Explain the steps to load an XLS file, build a PivotTable, add a column PivotChart, set series colors, and save as XLSX with Aspose.Cells.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

// Loads an XLS file (or generates sample data), builds a PivotTable on A1:C10, adds Quantity and Revenue as data fields, creates a linked column PivotChart, assigns blue to the Quantity series and green to the Revenue series, sets a chart title, and saves the result as an XLSX workbook.
class PivotChartWithColors
{
    static void Main()
    {
        try
        {
            const string inputFile = "InputData.xls";
            const string outputFile = "OutputWithPivotChart.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a sample workbook
            if (File.Exists(inputFile))
            {
                workbook = new Workbook(inputFile);
            }
            else
            {
                // Create a new workbook with sample data in A1:C10
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";

                // Header row
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["C1"].PutValue("Revenue");

                // Sample rows
                string[] products = { "A", "B", "C", "D", "E", "F", "G", "H", "I" };
                Random rnd = new Random();
                for (int i = 0; i < products.Length; i++)
                {
                    int row = i + 2;
                    sheet.Cells[row, 0].PutValue(products[i]);                     // Product
                    sheet.Cells[row, 1].PutValue(rnd.Next(10, 100));               // Quantity
                    sheet.Cells[row, 2].PutValue(rnd.Next(1000, 5000));            // Revenue
                }
            }

            // Use the first worksheet (adjust if needed)
            Worksheet dataSheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Create a PivotTable based on the source data
            // -------------------------------------------------
            // Assume source data is in range A1:C10 (adjust as per your file)
            int pivotIndex = dataSheet.PivotTables.Add("A1:C10", "E3", "SalesPivot");
            PivotTable pivotTable = dataSheet.PivotTables[pivotIndex];

            // Add a row field (e.g., Product)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column A

            // Add two value fields (e.g., Quantity and Revenue)
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Column B
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2); // Column C

            // Calculate the pivot data so the chart can read it
            pivotTable.CalculateData();

            // -------------------------------------------------
            // 2. Create a PivotChart linked to the PivotTable
            // -------------------------------------------------
            // Add a column chart (you can choose other types)
            int chartIndex = dataSheet.Charts.Add(ChartType.Column, 15, 0, 30, 15);
            Chart chart = dataSheet.Charts[chartIndex];

            // Link the chart to the pivot table
            chart.PivotSource = $"{dataSheet.Name}!SalesPivot";

            // Refresh chart data from the pivot table
            chart.RefreshPivotData();

            // -------------------------------------------------
            // 3. Assign individual colors to each value series
            // -------------------------------------------------
            // The first series corresponds to the first data field (Quantity)
            if (chart.NSeries.Count > 0)
                chart.NSeries[0].Area.ForegroundColor = Color.Blue;

            // The second series corresponds to the second data field (Revenue)
            if (chart.NSeries.Count > 1)
                chart.NSeries[1].Area.ForegroundColor = Color.Green;

            // Optional: set chart title
            chart.Title.Text = "Sales Overview";

            // -------------------------------------------------
            // 4. Save the workbook with the new PivotChart
            // -------------------------------------------------
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
