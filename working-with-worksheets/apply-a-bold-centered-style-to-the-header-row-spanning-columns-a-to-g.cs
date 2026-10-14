// Title: How to apply a bold, centered style to the header row across columns A‑G using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that defines a bold, center‑aligned style and uses Aspose.Cells ApplyStyle to format the range A1:G1. | Show how to create a Style and StyleFlag in Aspose.Cells and apply them to the first worksheet row spanning columns A through G.
// Common Searches: Aspose.Cells C# set bold centered formatting for header row A1:G1 | C# apply style to multiple columns in Excel using Aspose.Cells | How to use StyleFlag to format a range of cells in Aspose.Cells .NET | Create a bold header across columns A to G with Aspose.Cells example
// Tags: Aspose.Cells StyleFlag apply style to range | C# bold centered header formatting Aspose.Cells | Excel workbook header row styling Aspose.Cells | CreateStyle ApplyStyle C# example | A1:G1 range styling Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// The program creates a new workbook, defines a bold and center‑aligned style with StyleFlag, applies it to the header range A1:G1 on the first worksheet, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a style with bold font and centered alignment
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.HorizontalAlignment = TextAlignmentType.Center;

            // Define which style attributes to apply
            StyleFlag styleFlag = new StyleFlag
            {
                FontBold = true,
                HorizontalAlignment = true
            };

            // Apply the style to the header row spanning columns A to G (A1:G1)
            AsposeRange headerRange = sheet.Cells.CreateRange("A1:G1");
            headerRange.ApplyStyle(headerStyle, styleFlag);

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
