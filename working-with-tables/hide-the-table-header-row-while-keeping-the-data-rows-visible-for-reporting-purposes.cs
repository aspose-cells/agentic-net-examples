// Title: Hide the first table header row in an Excel worksheet while keeping data rows visible using Aspose.Cells for .NET
// AI Prompts: Hide the header row (row index 0) of a worksheet in an existing Excel file with Aspose.Cells for .NET. | Adjust the example to hide a consecutive range of rows (e.g., rows 0‑2) and leave all other rows displayed. | Save the modified workbook to a new file path after hiding the specified rows using Aspose.Cells.
// Common Searches: Aspose.Cells C# hide first row of Excel sheet without deleting data | How to make Excel header row invisible using Aspose.Cells for .NET | Programmatically hide table header in .xlsx while keeping rows visible with Aspose.Cells | C# set row IsHidden property in Aspose.Cells example | Hide multiple rows in an Excel workbook using Aspose.Cells C#
// Tags: hide worksheet row Aspose.Cells | Excel row IsHidden property C# | Aspose.Cells hide header row | preserve data rows while hiding header | save workbook with hidden rows Aspose.Cells | modify existing Excel file Aspose.Cells .NET

using Aspose.Cells;
using System;

// Loads input.xlsx, sets the first row's IsHidden flag to true to hide the header, and saves the result as output.xlsx while leaving all data rows visible.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index or name as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Hide the header row (row index 0). Data rows remain visible.
        worksheet.Cells.Rows[0].IsHidden = true;

        // Save the workbook with the hidden header row
        workbook.Save("output.xlsx");
    }
}
