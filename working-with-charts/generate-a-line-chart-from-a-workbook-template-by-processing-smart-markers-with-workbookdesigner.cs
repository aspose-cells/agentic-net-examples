// Title: Create a line chart in an Excel template after processing Smart Markers with WorkbookDesigner using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an existing Excel template, binds a DataTable to WorkbookDesigner, processes Smart Markers, and inserts a line chart referencing the populated range. | Modify the example to produce a multi‑series column chart by adding extra columns to the DataTable and mapping each column to a separate series. | Extend the program to save the generated chart as a PNG image file while still preserving the workbook.
// Common Searches: Aspose.Cells how to add a line chart after smart marker processing in C# | C# generate Excel chart from DataTable using WorkbookDesigner | Create chart in Excel template with Aspose.Cells .NET example | Smart markers populate data then create chart Aspose.Cells tutorial | Export Aspose.Cells chart to image file after workbook save
// Tags: Aspose.Cells WorkbookDesigner smart markers | Aspose.Cells add line chart programmatically | Aspose.Cells chart from DataTable range | Aspose.Cells export chart as image | Aspose.Cells generate chart in Excel template

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

// Loads an Excel template, fills it with a DataTable via WorkbookDesigner, processes Smart Markers, adds a line chart that references the populated cells, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Ensure the template file exists; create a simple one if missing
            if (!File.Exists(templatePath))
            {
                var tempWb = new Workbook();
                var tempSheet = tempWb.Worksheets[0];
                tempSheet.Cells[0, 0].PutValue("Month");
                tempSheet.Cells[0, 1].PutValue("Value");
                tempWb.Save(templatePath);
            }

            // Load the workbook template that contains Smart Markers
            Workbook workbook = new Workbook(templatePath);

            // Prepare the data source for Smart Markers
            DataTable dt = new DataTable("SalesData");
            dt.Columns.Add("Month", typeof(string));
            dt.Columns.Add("Value", typeof(double));
            dt.Rows.Add("Jan", 120);
            dt.Rows.Add("Feb", 150);
            dt.Rows.Add("Mar", 130);
            dt.Rows.Add("Apr", 170);
            dt.Rows.Add("May", 160);

            // Set up the WorkbookDesigner with the data source
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource(dt);
            designer.Process(); // Process Smart Markers

            // After processing, add a line chart to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a line chart at row 5, column 0 with size spanning rows 5-20 and columns 0-10
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Monthly Sales";

            // Define the data range (A2:B6) after Smart Marker processing
            int firstRow = 1;   // zero‑based index for row 2
            int firstCol = 0;   // column A
            int lastRow = 5;    // zero‑based index for row 6
            int monthCol = firstCol;       // column A contains Month
            int valueCol = firstCol + 1;   // column B contains Value

            // Create ranges for Y‑values and X‑values
            AsposeRange yRange = sheet.Cells.CreateRange(firstRow, valueCol, lastRow - firstRow + 1, 1);
            AsposeRange xRange = sheet.Cells.CreateRange(firstRow, monthCol, lastRow - firstRow + 1, 1);

            // Add the Y‑values series using the range reference string
            chart.NSeries.Add(yRange.RefersTo, true);
            // Set the X‑values (Month) for the series
            chart.NSeries[0].XValues = xRange.RefersTo;
            chart.NSeries[0].Name = "Sales";

            // Save the workbook with the generated chart
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
