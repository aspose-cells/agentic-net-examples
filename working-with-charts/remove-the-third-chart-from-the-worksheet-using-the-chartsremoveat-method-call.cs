// Title: C# – Remove the third chart from a worksheet using Aspose.Cells Charts.RemoveAt
// Description: Creates a workbook, adds column, line, and pie charts, then deletes the third chart (index 2) with sheet.Charts.RemoveAt(2) and saves the file as RemovedThirdChart.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells remove chart C# | Charts.RemoveAt example | delete specific chart worksheet | Aspose.Cells chart manipulation | remove third chart Aspose | C# Excel chart removal
// Common Searches: how to delete a chart in Aspose.Cells C# | Aspose.Cells Charts.RemoveAt usage | remove chart by index Aspose.Cells | C# code to delete third chart in Excel workbook | Aspose.Cells chart collection remove
// Developer Intent: Delete the chart at index 2 from a worksheet's Charts collection.
// Use Cases: Generate several charts for a report, then discard placeholder visuals before publishing. | Conditionally eliminate a chart based on user selection or business rules. | Clean up automatically added charts in a template to retain only required graphics.
// AI Prompts: Provide C# code that removes a chart at a given zero‑based index using Aspose.Cells Charts.RemoveAt. | Explain how to check the remaining number of charts after calling RemoveAt on a worksheet. | Suggest ways to handle IndexOutOfRangeException when the specified chart index does not exist.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds column, line, and pie charts, then deletes the third chart (index 2) with sheet.Charts.RemoveAt(2) and saves the file as RemovedThirdChart.xlsx using Aspose.Cells for .NET.
class RemoveThirdChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the charts
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
        sheet.Charts[chartIndex1].NSeries.Add("B2:B4", true);
        sheet.Charts[chartIndex1].NSeries.CategoryData = "A2:A4";

        int chartIndex2 = sheet.Charts.Add(ChartType.Line, 16, 0, 26, 8);
        sheet.Charts[chartIndex2].NSeries.Add("B2:B4", true);
        sheet.Charts[chartIndex2].NSeries.CategoryData = "A2:A4";

        int chartIndex3 = sheet.Charts.Add(ChartType.Pie, 27, 0, 37, 8);
        sheet.Charts[chartIndex3].NSeries.Add("B2:B4", true);
        sheet.Charts[chartIndex3].NSeries.CategoryData = "A2:A4";

        // Remove the third chart (zero‑based index 2)
        sheet.Charts.RemoveAt(2);

        // Save the workbook
        workbook.Save("RemovedThirdChart.xlsx");
    }
}
