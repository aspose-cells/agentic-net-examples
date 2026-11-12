// Title: Show blank cells as "N/A" with a custom number format in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# custom number format to show N/A for blank cells | how to apply a style to the entire used range in an Aspose.Cells workbook | replace null values with literal text in Excel using Aspose.Cells .NET | display placeholder text for empty cells in generated XLSX with Aspose.Cells
// Tags: custom number format empty cells Aspose.Cells | style applied to worksheet used range Aspose.Cells | null cell placeholder N/A Excel .NET | styleflag numberformat true Aspose.Cells | display N/A for blank cells Excel

using Aspose.Cells;
using System;

// The example creates a workbook, inserts numeric and null values, defines a custom number format that renders empty cells as "N/A", applies this style to the worksheet's used range via a StyleFlag, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data, including null (empty) cells
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["A3"].PutValue(null); // empty cell
            sheet.Cells["A4"].PutValue(20);
            sheet.Cells["A5"].PutValue(null); // empty cell

            // Create a style with a custom number format:
            // Positive;Negative;Zero;Text (fourth part handles empty/text cells)
            // The fourth part displays the literal string N/A for empty cells
            Style naStyle = workbook.CreateStyle();
            naStyle.Custom = "General;General;General;\"N/A\"";

            // Apply the style to the used range of the worksheet
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
            StyleFlag flag = new StyleFlag
            {
                NumberFormat = true // Apply only the custom number format
            };

            usedRange.ApplyStyle(naStyle, flag);

            // Save the workbook
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
