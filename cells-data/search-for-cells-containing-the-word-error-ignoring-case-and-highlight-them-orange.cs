// Title: Highlight cells containing the word “error” (case‑insensitive) with an orange background using Aspose.Cells for .NET (C#)
// AI Prompts: Find all cells whose value contains the substring "error" without case sensitivity and set an orange solid fill style using Aspose.Cells in C#. | Configure FindOptions for a case‑insensitive, contains search and apply a custom orange Style to each matching cell in a worksheet.
// Common Searches: aspocells c# find cells containing text case insensitive | how to highlight cells with specific word in Excel using Aspose.Cells | apply orange background to cells that contain 'error' in C# Aspose.Cells | using FindOptions LookAt Contains to style matching cells in Aspose.Cells
// Tags: Aspose.Cells case-insensitive FindOptions | orange cell background style Aspose.Cells | highlight cells containing substring Aspose.Cells | C# Excel cell search and style Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// Creates or loads a workbook, configures a case‑insensitive FindOptions with LookAt.Contains, iterates through all cells that contain the word "error", applies an orange solid fill style, and saves the result as HighlightedErrors.xlsx.
class HighlightErrorCells
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Sample data – this section can be removed when using a real workbook
        worksheet.Cells["A1"].PutValue("No error here");
        worksheet.Cells["A2"].PutValue("Error occurred");
        worksheet.Cells["A3"].PutValue("Another ERROR found");
        worksheet.Cells["A4"].PutValue("All good");

        // Configure find options: case‑insensitive, search cell values, match if contains the text
        FindOptions findOptions = new FindOptions
        {
            CaseSensitive = false,
            LookInType = LookInType.Values,
            LookAtType = LookAtType.Contains
        };

        // Create a style with orange background
        Style orangeStyle = workbook.CreateStyle();
        orangeStyle.ForegroundColor = Color.Orange;
        orangeStyle.Pattern = BackgroundType.Solid;

        // Iterate through all cells that contain the word "error" (ignoring case) and apply the style
        Cell previousCell = null;
        Cell foundCell = worksheet.Cells.Find("error", previousCell, findOptions);
        while (foundCell != null)
        {
            foundCell.SetStyle(orangeStyle);
            previousCell = foundCell;
            foundCell = worksheet.Cells.Find("error", previousCell, findOptions);
        }

        // Save the modified workbook
        workbook.Save("HighlightedErrors.xlsx");
    }
}
