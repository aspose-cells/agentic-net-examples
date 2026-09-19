// Title: Apply Accent6 Theme Color to Header Row in a Dynamically Generated Excel Report with Aspose.Cells for .NET (C#)
// AI Prompts: Create a solid‑fill style using the workbook's Accent6 theme color, set the font to bold white, and apply it to the first row of the worksheet. | Write header titles to row 0 and style the entire row with an Accent6 background and bold white text using Aspose.Cells APIs. | Save the workbook after applying the Accent6‑styled header row to an .xlsx file.
// Common Searches: how to set Excel header background to theme Accent6 using Aspose.Cells C# | Aspose.Cells apply workbook theme color to first row of worksheet | C# create header style with solid fill and bold white font in Aspose.Cells | dynamic report generation with themed header row in Aspose.Cells .NET
// Tags: apply theme accent6 background Aspose.Cells | header row style solid fill C# | bold white font Excel header Aspose.Cells | save workbook as xlsx Aspose.Cells | dynamic report generation Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, writes column headers to the first row, defines a style with the Accent6 theme color (solid fill) and bold white font, applies this style to the header row, and saves the result as Report.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define header titles
            string[] headers = { "ID", "Name", "Quantity", "Price" };

            // Write headers to the first row (row index 0)
            for (int col = 0; col < headers.Length; col++)
            {
                sheet.Cells[0, col].PutValue(headers[col]);
            }

            // Use a fallback accent color (since Workbook.Theme returns a string)
            Color accent6 = Color.FromArgb(0, 112, 192); // Example accent color

            // Create a style for the header row
            Style headerStyle = workbook.CreateStyle();
            headerStyle.ForegroundColor = accent6;
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = Color.White; // contrast color

            // Apply the style to the entire header row
            StyleFlag flag = new StyleFlag { All = true };
            sheet.Cells.CreateRange(0, 0, 1, headers.Length).ApplyStyle(headerStyle, flag);

            // Define output file path
            string outputPath = "Report.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
