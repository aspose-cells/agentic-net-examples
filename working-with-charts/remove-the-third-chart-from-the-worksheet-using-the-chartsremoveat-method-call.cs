// Title: How to delete the third chart from an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, verifies that at least three charts exist on the first worksheet, removes the chart at zero‑based index 2, and saves the workbook. | Generate a C# snippet that calls Worksheet.Charts.RemoveAt to delete a chart by its position and includes logic to handle insufficient chart count.
// Common Searches: Aspose.Cells C# remove chart at index 2 from a worksheet | how to delete the third chart in an Excel file using Aspose.Cells | using Worksheet.Charts.RemoveAt to erase a specific chart in .NET | C# example for checking chart count before removing a chart with Aspose.Cells | save Excel workbook after removing a chart with Aspose.Cells for .NET
// Tags: Worksheet.Charts.RemoveAt method Aspose.Cells | delete chart by zero‑based index C# Aspose.Cells | chart count validation before removal Aspose.Cells | save workbook after chart deletion Aspose.Cells | C# Excel chart manipulation Aspose.Cells

using Aspose.Cells;

// The sample loads 'input.xlsx', accesses the first worksheet, ensures there are at least three charts, removes the chart at zero‑based index 2 using Worksheet.Charts.RemoveAt, and saves the modified file as 'output.xlsx'.
class Program
{
    static void Main()
    {
        // Load the workbook that contains the charts
        Workbook workbook = new Workbook("input.xlsx");

        // Access the target worksheet (first sheet in this example)
        Worksheet sheet = workbook.Worksheets[0];

        // Remove the third chart (index is zero‑based, so index 2)
        if (sheet.Charts.Count > 2)
        {
            sheet.Charts.RemoveAt(2);
        }

        // Save the workbook after the chart has been removed
        workbook.Save("output.xlsx");
    }
}
