// Title: Add a conditional IF smart marker in Aspose.Cells for .NET to label sales as High or Low using a threshold parameter
// AI Prompts: Create a smart marker expression that uses the IF function to compare each sales cell with a Threshold column from a DataTable and returns "High" or "Low", then invoke WorkbookDesigner.Process to generate the final Excel workbook. | Place the smart marker in cell B2, bind a DataTable containing the threshold value, process the markers with WorkbookDesigner, and save the file as SalesConditionalSmartMarker.xlsx using Aspose.Cells for .NET.
// Common Searches: aspnet conditional smart marker with if statement in Aspose.Cells | how to bind a DataTable threshold to a smart marker formula in Aspose.Cells | create high low labels in Excel using Aspose.Cells smart markers | using WorkbookDesigner to process IF smart markers based on a data source | Aspose.Cells smart marker syntax for conditional labeling
// Tags: IF smart marker expression Aspose.Cells | WorkbookDesigner data source threshold | conditional labeling with smart markers Excel | smart marker formula using DataTable | Aspose.Cells generate high low sales report

using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsSmartMarkerExample
{
    // The example creates a workbook, writes sample sales values, inserts a conditional smart marker that evaluates each sales entry against a threshold supplied via a DataTable, processes the marker with WorkbookDesigner, and saves the result as SalesConditionalSmartMarker.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add header for sales data
                sheet.Cells["A1"].PutValue("Sales");

                // Sample sales values
                double[] salesData = { 80, 120, 95, 150 };
                for (int i = 0; i < salesData.Length; i++)
                {
                    // Column A (0-index) rows start at 1 (A2, A3, ...)
                    sheet.Cells[i + 1, 0].PutValue(salesData[i]);
                }

                // Insert a smart marker that evaluates the sales value against a threshold
                // The smart marker syntax: &=IF(A2>{Threshold},"High","Low")
                // It will be copied down automatically when processed
                sheet.Cells["B2"].PutValue("&=IF(A2>{Threshold},\"High\",\"Low\")");

                // Prepare a data source containing the threshold value
                DataTable dt = new DataTable("Parameters");
                dt.Columns.Add("Threshold", typeof(double));
                dt.Rows.Add(100); // Example threshold

                // Set up the WorkbookDesigner, assign the data source, and process the smart markers
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                designer.SetDataSource(dt);
                designer.Process();

                // Save the resulting workbook
                string outputPath = "SalesConditionalSmartMarker.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
