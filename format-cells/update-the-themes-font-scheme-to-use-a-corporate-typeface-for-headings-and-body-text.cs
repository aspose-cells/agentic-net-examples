// Title: Set Corporate Heading and Body Fonts in an Excel Workbook with Aspose.Cells for .NET
// Description: Demonstrates how to create a new workbook, define corporate heading and body typefaces, apply a 20‑pt bold heading font to cell A1 and a 12‑pt body font to cell A3, and save the file using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# Excel font styling | custom corporate font | heading font Aspose.Cells | body text font .NET | programmatic Excel formatting | Excel workbook theme fonts | apply font to cells | enterprise Excel branding
// Common Searches: how to set custom heading font in Aspose.Cells C# | apply corporate body font to Excel cells programmatically | change font of specific cells using Aspose.Cells .NET | set theme fonts for Excel workbook with Aspose | C# code to use corporate typeface in Excel
// Developer Intent: Apply corporate typefaces to heading and body cells in an Excel workbook via Aspose.Cells for .NET.
// Use Cases: Generate branded reports where the title uses a corporate heading font and the content uses the company’s body font. | Create a reusable Excel template that automatically formats headings and body text with specified corporate fonts. | Automate the styling of financial statements to match corporate brand guidelines.
// AI Prompts: Show me how to modify an Excel workbook’s theme font scheme globally to use corporate heading and body fonts with Aspose.Cells for .NET. | Provide C# code that loads an existing .xlsx file and replaces its theme’s heading and body fonts with specified corporate typefaces. | Explain how to define a reusable style in Aspose.Cells that applies corporate fonts to multiple cells without setting each cell individually.

using System;
using System.IO;
using Aspose.Cells;

namespace CorporateThemeFontDemo
{
    // Demonstrates how to create a new workbook, define corporate heading and body typefaces, apply a 20‑pt bold heading font to cell A1 and a 12‑pt body font to cell A3, and save the file using Aspose.Cells for C#.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Define corporate typefaces (replace with actual font names)
                string corporateHeadingFont = "CorporateHeadingFont";
                string corporateBodyFont = "CorporateBodyFont";

                // -------------------------------------------------
                // Apply corporate heading font to a heading cell
                Cell headingCell = sheet.Cells["A1"];
                headingCell.PutValue("Report Title");

                // Get the style of the heading cell and modify its font
                Style headingStyle = headingCell.GetStyle();
                headingStyle.Font.Name = corporateHeadingFont;
                headingStyle.Font.Size = 20;
                headingStyle.Font.IsBold = true;

                // Assign the modified style back to the cell
                headingCell.SetStyle(headingStyle);

                // -------------------------------------------------
                // Apply corporate body font to a body text cell
                Cell bodyCell = sheet.Cells["A3"];
                bodyCell.PutValue("This is the body text of the report.");

                // Get the style of the body cell and modify its font
                Style bodyStyle = bodyCell.GetStyle();
                bodyStyle.Font.Name = corporateBodyFont;
                bodyStyle.Font.Size = 12;

                // Assign the modified style back to the cell
                bodyCell.SetStyle(bodyStyle);

                // -------------------------------------------------
                // Save the workbook
                string outputPath = "CorporateThemeFontDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
