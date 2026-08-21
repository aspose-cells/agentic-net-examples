// Title: Set Excel Theme Font Scheme (Major & Minor) to a Corporate Typeface with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, assign a corporate font (e.g., Calibri) to the Major (heading) and Minor (body) theme schemes using Font.SetName with FontSchemeType, apply the styles to cells, and save the file as CorporateThemeFontScheme.xlsx.
// Keywords: Aspose.Cells C# theme font scheme | FontSchemeType Major Minor | set corporate font Excel workbook | apply custom font to headings and body | Excel theme fonts Aspose.Cells .NET
// Common Searches: how to change major and minor font scheme in Aspose.Cells | set corporate typeface for Excel headings C# | apply custom theme fonts to Excel workbook using Aspose.Cells | Aspose.Cells FontSchemeType example | update Excel theme fonts programmatically .NET
// Developer Intent: Configure the workbook’s theme so that headings use the Major font scheme and body text uses the Minor scheme, both set to a corporate typeface.
// Use Cases: Create a new workbook and define a corporate font for the Major scheme to style report titles. | Apply the same corporate font to the Minor scheme for body paragraphs and data cells. | Save the workbook so that any subsequent styles inherit the corporate typeface automatically.
// AI Prompts: Generate C# code with Aspose.Cells that changes the Major and Minor theme fonts of an existing workbook to 'Arial' and saves the result. | Explain FontSchemeType in Aspose.Cells and show how to use it for heading and body styles. | Provide a step‑by‑step tutorial for updating an Excel workbook’s theme fonts to a corporate typeface using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, assign a corporate font (e.g., Calibri) to the Major (heading) and Minor (body) theme schemes using Font.SetName with FontSchemeType, apply the styles to cells, and save the file as CorporateThemeFontScheme.xlsx.
class UpdateThemeFontScheme
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Corporate typeface to be used for both headings (Major) and body text (Minor)
        string corporateFontName = "Calibri";

        // ---------- Heading style (Major scheme) ----------
        Style headingStyle = workbook.CreateStyle();
        // Set the font name and associate it with the Major scheme
        headingStyle.Font.SetName(corporateFontName, FontSchemeType.Major);
        headingStyle.Font.Size = 16;
        headingStyle.Font.IsBold = true;

        // Apply heading style to a cell
        worksheet.Cells["A1"].PutValue("Report Title");
        worksheet.Cells["A1"].SetStyle(headingStyle);

        // ---------- Body style (Minor scheme) ----------
        Style bodyStyle = workbook.CreateStyle();
        // Set the font name and associate it with the Minor scheme
        bodyStyle.Font.SetName(corporateFontName, FontSchemeType.Minor);
        bodyStyle.Font.Size = 11;

        // Apply body style to a cell
        worksheet.Cells["A2"].PutValue("This is body text using the corporate typeface.");
        worksheet.Cells["A2"].SetStyle(bodyStyle);

        // Save the workbook
        workbook.Save("CorporateThemeFontScheme.xlsx");
    }
}
