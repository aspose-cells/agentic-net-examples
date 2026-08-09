// Title: Aspose.Cells for .NET – Remove the third chart using Charts.RemoveAt (C#)
// Description: Demonstrates how to create a workbook, add three charts (column, line, pie), display the chart count, delete the third chart with a zero‑based index (2) via sheet.Charts.RemoveAt(2), verify the updated count, and save the file as RemoveThirdChartDemo.xlsx.
// Keywords: Aspose.Cells remove chart C# | Charts.RemoveAt example | delete specific chart Aspose.Cells | remove third chart worksheet | Aspose.Cells chart management .NET | C# Aspose.Cells chart deletion
// Common Searches: How to delete a chart by index in Aspose.Cells C# | Remove third chart from worksheet Aspose.Cells | Charts.RemoveAt usage Aspose.Cells .NET | Aspose.Cells delete specific chart example | C# code to remove a chart from Excel workbook
// Developer Intent: Programmatically delete the third chart in a worksheet.
// Use Cases: Eliminate an automatically generated placeholder chart before publishing a report. | Adjust the visual layout of a workbook by removing unwanted charts after dynamic creation. | Maintain a clean worksheet when the number of charts varies based on user input.
// AI Prompts: Show a C# snippet that removes the third chart from an Aspose.Cells worksheet using Charts.RemoveAt. | Explain how chart indices shift after calling Charts.RemoveAt in Aspose.Cells. | Provide a verification step to confirm chart removal by checking the Charts.Count property.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add three charts (column, line, pie), display the chart count, delete the third chart with a zero‑based index (2) via sheet.Charts.RemoveAt(2), verify the updated count, and save the file as RemoveThirdChartDemo.xlsx.
class RemoveThirdChartDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the charts
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add three charts to the worksheet
        int chartIndex1 = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 8);
        int chartIndex2 = sheet.Charts.Add(ChartType.Line, 16, 0, 26, 8);
        int chartIndex3 = sheet.Charts.Add(ChartType.Pie, 27, 0, 37, 8);

        // Configure each chart (optional, just to have valid data)
        sheet.Charts[chartIndex1].NSeries.Add("B2:B4", true);
        sheet.Charts[chartIndex1].NSeries.CategoryData = "A2:A4";

        sheet.Charts[chartIndex2].NSeries.Add("B2:B4", true);
        sheet.Charts[chartIndex2].NSeries.CategoryData = "A2:A4";

        sheet.Charts[chartIndex3].NSeries.Add("B2:B4", true);
        sheet.Charts[chartIndex3].NSeries.CategoryData = "A2:A4";

        // Display count before removal
        Console.WriteLine("Chart count before removal: " + sheet.Charts.Count);

        // Remove the third chart (zero‑based index = 2)
        sheet.Charts.RemoveAt(2);

        // Display count after removal
        Console.WriteLine("Chart count after removal: " + sheet.Charts.Count);

        // Save the workbook
        workbook.Save("RemoveThirdChartDemo.xlsx");
    }
}
