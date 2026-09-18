// Title: Add conditional formatting to a cell range and retain the colors when saving the workbook as PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Create a conditional formatting rule that colors cells light green when the value exceeds 50 for range A1:A10, then export the workbook to PDF preserving the formatting. | Define a conditional formatting area, set a greater‑than‑50 condition with a solid light‑green background, and save the Excel file as a PDF using Aspose.Cells in C#. | Programmatically apply a cell‑value based conditional format and generate a PDF that shows the conditional colors with the Aspose.Cells .NET API.
// Common Searches: Aspose.Cells C# how to keep conditional formatting colors in PDF output | example of adding conditional formatting to Excel and exporting to PDF with Aspose.Cells .NET | preserve conditional formatting background when converting workbook to PDF using Aspose.Cells | C# code to apply greater than condition and export to PDF with Aspose.Cells
// Tags: conditional formatting setup Aspose.Cells C# | PDF export with retained formatting Aspose.Cells | greater‑than value condition Aspose.Cells | cell background style configuration Aspose.Cells | dynamic conditional‑formatting handling Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The program loads or creates an Excel workbook, defines a conditional formatting rule that colors cells light green when their value exceeds 50 for the range A1:A10, and then saves the workbook as a PDF, preserving the conditional formatting colors.
class Program
{
    static void Main()
    {
        try
        {
            // Load the source Excel workbook; create a new one if the file does not exist.
            Workbook workbook;
            const string inputPath = "input.xlsx";

            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Define the range for conditional formatting (e.g., A1:A10)
            int firstRow = 0;          // Row index for A1
            int firstColumn = 0;       // Column index for A1
            int totalRows = 10;        // Number of rows in the range
            int totalColumns = 1;      // Number of columns in the range

            // Add a new conditional formatting rule to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();

            // Use dynamic to avoid compile‑time dependency on the ConditionalFormatting type
            dynamic conditionalFormatting = sheet.ConditionalFormattings[cfIndex];

            // Set the target area for the conditional formatting
            conditionalFormatting.AddArea(firstRow, firstColumn, totalRows, totalColumns);

            // Create a condition: cell value greater than 50
            int conditionIndex = conditionalFormatting.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "50");

            // Retrieve the condition and define its style
            dynamic condition = conditionalFormatting[conditionIndex];
            dynamic style = condition.Style;
            style.ForegroundColor = Color.LightGreen;
            style.Pattern = BackgroundType.Solid;

            // Save the workbook as PDF; conditional formatting colors will be reflected in the PDF
            const string outputPath = "output.pdf";
            workbook.Save(outputPath, SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
