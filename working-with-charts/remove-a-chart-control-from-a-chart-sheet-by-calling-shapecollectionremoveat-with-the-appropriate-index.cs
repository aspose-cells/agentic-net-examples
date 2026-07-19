// Title: Aspose.Cells .NET – Delete a Chart Shape from a Worksheet Using ShapeCollection.RemoveAt
// Description: Demonstrates how to create a workbook, add a column chart, populate it with data, and then remove the chart (which is stored as a shape) from the worksheet by calling ShapeCollection.RemoveAt with the correct index, finally saving the file as ChartSheetWithoutChart.xlsx.
// Keywords: Aspose.Cells remove chart shape | ShapeCollection.RemoveAt example | delete chart worksheet .NET | Aspose.Cells chart removal | C# Excel chart shape delete | Aspose.Cells ShapeCollection | remove Excel chart programmatically
// Common Searches: how to delete a chart in Aspose.Cells | remove chart shape from Excel file using C# | ShapeCollection.RemoveAt chart index | Aspose.Cells delete chart programmatically | C# remove Excel chart object
// Developer Intent: The developer needs to programmatically eliminate a chart object from a worksheet by invoking ShapeCollection.RemoveAt with the appropriate shape index.
// Use Cases: Erase a temporary chart after exporting data to keep the workbook clean. | Clear pre‑existing chart placeholders in a template before inserting new visualizations. | Strip unwanted chart objects from a workbook prior to distribution or archiving.
// AI Prompts: Generate C# code that removes a specific chart shape from an Aspose.Cells worksheet using ShapeCollection.RemoveAt. | Explain how to locate the index of a chart shape when multiple shapes exist on a worksheet. | Show a loop that iterates through a worksheet's Shapes collection and safely deletes only chart shapes.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a column chart, populate it with data, and then remove the chart (which is stored as a shape) from the worksheet by calling ShapeCollection.RemoveAt with the correct index, finally saving the file as ChartSheetWithoutChart.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a regular worksheet that will hold the chart
                Worksheet sheet = workbook.Worksheets.Add("MyChartSheet");

                // Add a chart to the worksheet (the chart is also a shape in the sheet)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Populate some data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                chart.NSeries.Add("B2:B3", true);
                chart.NSeries.CategoryData = "A2:A3";

                // Remove the chart control (shape) from the worksheet using ShapeCollection.RemoveAt
                // Assuming the chart is the first (and only) shape in the collection
                if (sheet.Shapes.Count > 0)
                {
                    sheet.Shapes.RemoveAt(0);
                }

                // Define output file path
                string outputPath = "ChartSheetWithoutChart.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
