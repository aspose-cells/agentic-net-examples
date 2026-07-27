// Title: Save a modified Excel workbook with original formatting and textures using Aspose.Cells for .NET
// Description: Load an existing XLSX file, change cell values, and save the workbook to a new file while keeping all styles, textures, conditional formatting, images, and other visual elements unchanged.
// Keywords: Aspose.Cells save workbook preserve formatting | C# save edited Excel without losing styles | retain textures Aspose.Cells | keep original Excel design .NET | save workbook new file preserving visual elements
// Common Searches: Aspose.Cells save workbook keep original formatting | C# preserve Excel styles when saving with Aspose | how to retain textures in saved Excel file Aspose.Cells | save modified XLSX without losing conditional formatting .NET | duplicate Excel template and keep design using Aspose
// Developer Intent: Save an edited workbook to a separate file while preserving every visual attribute of the original spreadsheet.
// Use Cases: Update a template workbook (e.g., change dates or totals) and export a copy that retains the template’s layout and branding. | Generate a report by programmatically altering data in an existing spreadsheet and saving it as a new file without affecting conditional formatting or cell styles. | Create a versioned backup of an Excel file, apply automated changes, and write the result to another location while keeping embedded images, charts, and textures intact.
// AI Prompts: Write C# code that opens an Excel template with Aspose.Cells, modifies multiple cells, and saves the workbook to a new file while preserving all formatting, textures, and embedded objects. | Explain how to configure Aspose.Cells SaveOptions in .NET to ensure styles, conditional formatting, cell textures, and images are retained when saving a workbook as XLSX.

using System;
using Aspose.Cells;

// Load an existing XLSX file, change cell values, and save the workbook to a new file while keeping all styles, textures, conditional formatting, images, and other visual elements unchanged.
class Program
{
    static void Main()
    {
        // Load an existing workbook (preserves all original formatting, styles, textures, etc.)
        Workbook workbook = new Workbook("input.xlsx");

        // Example modification: change the value of cell A1 in the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Modified");

        // Save the modified workbook to a new file while keeping all original formatting intact
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
