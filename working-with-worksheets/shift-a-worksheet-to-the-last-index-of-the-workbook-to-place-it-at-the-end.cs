// Title: How to move a specific worksheet to the last tab of an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to load an existing .xlsx file, locate a worksheet by name, move it to the final index with Worksheet.MoveTo, and save the workbook. | Write C# code that reorders worksheets so the selected sheet becomes the last sheet in the workbook using the Aspose.Cells API. | Demonstrate programmatically shifting any worksheet to the end of an Excel workbook with Aspose.Cells for .NET.
// Common Searches: aspnet move worksheet to last position Aspose.Cells C# example | C# Aspose.Cells reorder sheets so selected sheet is last | how to place a specific Excel tab at the end using Aspose.Cells | move sheet to final index in workbook with Aspose.Cells .NET
// Tags: Worksheet.MoveTo usage Aspose.Cells | move worksheet to final tab C# | reorder workbook sheets Aspose.Cells .NET | place Excel sheet at end programmatically | Aspose.Cells sheet ordering example

using Aspose.Cells;

// Loads 'input.xlsx', finds the worksheet named 'Sheet1', moves it to the last index using Worksheet.MoveTo, and saves the updated workbook as 'output.xlsx' with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Retrieve the worksheet you want to move (by name or index)
        Worksheet sheet = workbook.Worksheets["Sheet1"]; // replace with your sheet name

        // Move the worksheet to the last position in the workbook
        sheet.MoveTo(workbook.Worksheets.Count - 1);

        // Save the workbook with the worksheet now at the end
        workbook.Save("output.xlsx");
    }
}
