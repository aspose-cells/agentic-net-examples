// Title: Bold & Center Header Row (A‑G) with Aspose.Cells for .NET
// Description: Creates a new Workbook, defines a style with bold font and centered alignment, applies it to the range A1:G1 using a StyleFlag, inserts sample header text, and saves the file as HeaderStyleDemo.xlsx.
// Keywords: Aspose.Cells header style C# | bold centered header Aspose.Cells | apply style to range A1:G1 | StyleFlag formatting .NET | Excel header formatting Aspose | C# Aspose.Cells example
// Common Searches: Aspose.Cells make header row bold and centered | C# apply style to cells A1 to G1 Aspose | StyleFlag usage Aspose.Cells .NET | format Excel header row programmatically | Aspose.Cells header row styling example
// Developer Intent: Format the first worksheet row (A1:G1) as a bold, horizontally and vertically centered header.
// Use Cases: Generate reports with a visually distinct title row. | Standardize header appearance across multiple exported Excel sheets. | Create reusable header styles for automated data population.
// AI Prompts: Write C# code using Aspose.Cells to apply a bold, centered style to the header row spanning columns A to G and add sample headings. | Show how to use StyleFlag in Aspose.Cells to apply only font boldness and alignment to a specific range. | Explain how to define a reusable header style and apply it to several worksheets in a workbook with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing; // for alignment enums if needed

// Creates a new Workbook, defines a style with bold font and centered alignment, applies it to the range A1:G1 using a StyleFlag, inserts sample header text, and saves the file as HeaderStyleDemo.xlsx.
class ApplyHeaderStyle
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Create a style for the header: bold font and centered alignment
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;                         // make text bold
            headerStyle.HorizontalAlignment = TextAlignmentType.Center; // center horizontally
            headerStyle.VerticalAlignment = TextAlignmentType.Center;   // optional: center vertically

            // Define which style properties should be applied
            StyleFlag flag = new StyleFlag();
            flag.FontBold = true;               // apply bold setting
            flag.HorizontalAlignment = true;   // apply horizontal alignment
            flag.VerticalAlignment = true;      // apply vertical alignment (optional)

            // Create a range that covers columns A to G in the first row (row index 0)
            // Parameters: startRow, startColumn, totalRows, totalColumns
            Aspose.Cells.Range headerRange = cells.CreateRange(0, 0, 1, 7); // A1:G1

            // Apply the style to the defined range
            headerRange.ApplyStyle(headerStyle, flag);

            // Optionally, put some sample header text
            string[] headers = { "Header1", "Header2", "Header3", "Header4", "Header5", "Header6", "Header7" };
            for (int i = 0; i < headers.Length; i++)
            {
                cells[0, i].PutValue(headers[i]);
            }

            // Save the workbook
            workbook.Save("HeaderStyleDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
