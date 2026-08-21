// Title: Create a Bold, Center‑Aligned Header Across A‑G Using Aspose.Cells (C#)
// Description: Shows how to build a new Workbook, set up a Style with bold font and center alignment via StyleFlag, apply it to the first row covering columns A‑G, auto‑fit the row height, and save the file as HeaderStyle.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# style range | bold header Aspose.Cells | center‑aligned header row | StyleFlag font bold | HorizontalAlignment Center | AutoFitRow Excel | Excel header formatting .NET | global spreadsheet styling | US developers Aspose.Cells | India .NET Excel library
// Common Searches: Aspose.Cells how to make header row bold and centered | C# apply style to A-G range Aspose.Cells | Set horizontal alignment center with StyleFlag Aspose.Cells | AutoFitRow after styling header Aspose.Cells | Create bold centered title row in Excel using Aspose.Cells .NET | Aspose.Cells tutorial for header formatting
// Developer Intent: Apply bold font and center alignment to the first worksheet row spanning columns A‑G.
// Use Cases: Generate a report template with a prominent title row. | Prepare export files where the header spans multiple columns for better readability. | Automate financial dashboards that require a consistent header style. | Build multi‑regional spreadsheets (US, EU, APAC) with a unified header format.
// AI Prompts: Write C# code using Aspose.Cells to apply a bold, centered style to cells A1:G1 and auto‑fit the row. | Explain the steps to create a Style, configure a StyleFlag, and apply it to a range in Aspose.Cells for .NET. | Show how to encapsulate header styling into a reusable method that accepts any worksheet and column range. | Provide a GitHub‑ready snippet that formats a header row across A‑G with bold font and center alignment using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing; // for TextAlignmentType
using AsposeRange = Aspose.Cells.Range; // Resolve ambiguity with System.Range

// Shows how to build a new Workbook, set up a Style with bold font and center alignment via StyleFlag, apply it to the first row covering columns A‑G, auto‑fit the row height, and save the file as HeaderStyle.xlsx with Aspose.Cells for .NET.
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

            // Create a style: bold font and centered horizontally
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.HorizontalAlignment = TextAlignmentType.Center;

            // Define which style properties to apply
            StyleFlag flag = new StyleFlag
            {
                FontBold = true,
                HorizontalAlignment = true
            };

            // Create a range that covers columns A (0) to G (6) in the first row (row index 0)
            AsposeRange headerRange = cells.CreateRange(0, 0, 1, 7); // startRow, startColumn, totalRows, totalColumns

            // Apply the style to the defined range
            headerRange.ApplyStyle(headerStyle, flag);

            // Optionally autofit the row height to display the centered text nicely
            worksheet.AutoFitRow(0);

            // Save the workbook
            workbook.Save("HeaderStyle.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
