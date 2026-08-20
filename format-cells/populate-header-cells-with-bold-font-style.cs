// Title: Apply Bold Font to Header Row Using Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, insert header values, define a bold font style, limit the change with a StyleFlag, apply the style to the first row via ApplyRowStyle, and save the file as HeaderBold.xlsx.
// Keywords: Aspose.Cells C# bold header | ApplyRowStyle bold font | StyleFlag FontBold Aspose | Excel header formatting .NET | C# Aspose.Cells style row | Excel bold column headers
// Common Searches: Aspose.Cells make first row bold C# | C# apply bold style to Excel header Aspose | StyleFlag only bold font Aspose.Cells | How to format header row in Excel using Aspose.Cells .NET | Save workbook with bold headers Aspose
// Developer Intent: Add a bold font style to the worksheet’s header row.
// Use Cases: Produce sales reports where column titles stand out for quick scanning. | Export financial statements with emphasized headers to improve readability. | Create a reusable template that automatically formats header rows in bold for data entry forms.
// AI Prompts: Generate C# code that applies a bold font to multiple header rows with Aspose.Cells. | Explain how to use StyleFlag to change only the FontBold attribute in an Aspose.Cells style. | Provide an example of conditionally bolding header cells based on their text content using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsHeaderBoldDemo
{
    // Shows how to create a workbook, insert header values, define a bold font style, limit the change with a StyleFlag, apply the style to the first row via ApplyRowStyle, and save the file as HeaderBold.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate header cells (first row) with sample text
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Price");
            cells["C1"].PutValue("Quantity");

            // Create a style with bold font
            Style boldStyle = workbook.CreateStyle();
            boldStyle.Font.IsBold = true;

            // Create a StyleFlag that applies only the FontBold attribute
            StyleFlag flag = new StyleFlag { FontBold = true };

            // Apply the bold style to the entire first row (row index 0)
            worksheet.Cells.ApplyRowStyle(0, boldStyle, flag);

            // Save the workbook
            workbook.Save("HeaderBold.xlsx");

            Console.WriteLine("Header cells have been styled with bold font and saved to HeaderBold.xlsx");
        }
    }
}
