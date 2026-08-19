// Title: C# – Save a Modified Excel Workbook as XLSX While Retaining Original Formatting with Aspose.Cells
// Description: Load an existing .xlsx file, change cell values, and save it as a new workbook using Aspose.Cells for .NET. The SaveFormat.Xlsx option keeps all original styles, borders, number formats, formulas, and conditional formatting intact.
// Keywords: Aspose.Cells C# save workbook | preserve Excel formatting .NET | SaveFormat.Xlsx example | load and modify Excel file Aspose | retain cell styles when saving | C# Excel export without losing format
// Common Searches: Aspose.Cells save edited Excel without losing formatting | C# keep cell styles when saving workbook | How to preserve original Excel layout using Aspose.Cells | SaveFormat.Xlsx keep formulas and conditional formatting
// Developer Intent: Export a changed workbook to a new XLSX file without altering any of the source formatting.
// Use Cases: Update a template workbook and generate a styled report copy. | Apply data transformations across worksheets while preserving formulas and conditional rules. | Create versioned backups of an Excel file after programmatic edits, keeping the original design unchanged.
// AI Prompts: Generate C# code that loads an .xlsx file, modifies specific cells, and saves a new file with all original formatting using Aspose.Cells. | Explain step‑by‑step how SaveFormat.Xlsx maintains styles, borders, and number formats when saving a modified workbook. | Provide a tutorial for preserving formulas and conditional formatting while exporting a changed Excel workbook with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsSaveExample
{
    // Load an existing .xlsx file, change cell values, and save it as a new workbook using Aspose.Cells for .NET. The SaveFormat.Xlsx option keeps all original styles, borders, number formats, formulas, and conditional formatting intact.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (original formatting is retained)
            Workbook workbook = new Workbook("input.xlsx");

            // Example modification: update the value of cell A1
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("Modified");

            // Save the modified workbook to a new XLSX file while preserving formatting
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
