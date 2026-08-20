// Title: Set Global Workbook Date Format to dd-MMM-yyyy with Aspose.Cells for .NET (C#)
// Description: Shows how to modify a workbook's DefaultStyle in Aspose.Cells for .NET so every new cell automatically uses the custom date format dd-MMM-yyyy, then saves the workbook.
// Keywords: Aspose.Cells default date format C# | global Excel date style Aspose | custom date format dd-MMM-yyyy .NET | set workbook default style Aspose.Cells | C# Excel date formatting Aspose
// Common Searches: Aspose.Cells set default date format C# | How to apply a global date style in an Excel workbook using Aspose.Cells | Change default cell format to dd-MMM-yyyy with Aspose.Cells .NET | C# Aspose.Cells default style custom date format example
// Developer Intent: Configure the workbook’s default style so all cells inherit the dd-MMM-yyyy date format without per‑cell styling.
// Use Cases: Create a template where every date appears as dd-MMM-yyyy across all worksheets. | Generate periodic reports that require a consistent date appearance without manual formatting. | Load an existing workbook and enforce a new global date format before adding further data.
// AI Prompts: Write C# code using Aspose.Cells to set the workbook’s default date format to "dd-MMM-yyyy" and save the file. | Explain how the DefaultStyle.Custom property affects date formatting for newly added cells in Aspose.Cells. | Show how to apply a global custom date format and then override it for a specific column or range in an Excel workbook.

using System;
using Aspose.Cells;

// Shows how to modify a workbook's DefaultStyle in Aspose.Cells for .NET so every new cell automatically uses the custom date format dd-MMM-yyyy, then saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the workbook's default style to use the custom date format "dd-mmm-yyyy"
        Style defaultStyle = workbook.DefaultStyle;
        defaultStyle.Custom = "dd-mmm-yyyy";
        workbook.DefaultStyle = defaultStyle;

        // Optional: demonstrate the format by putting a date into a cell
        Worksheet sheet = workbook.Worksheets[0];
        Cell dateCell = sheet.Cells["A1"];
        dateCell.PutValue(DateTime.Now);
        // The cell will inherit the default style, so no need to set the style explicitly

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}
