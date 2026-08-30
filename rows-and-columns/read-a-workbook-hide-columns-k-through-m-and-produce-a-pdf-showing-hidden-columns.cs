// Title: Hide columns K‑M in an Excel worksheet and export to PDF while preserving hidden columns using Aspose.Cells for .NET
// AI Prompts: Hide columns K through M on the first worksheet and save the workbook as a PDF using Aspose.Cells in C#. | Apply the HideColumns method to conceal a range of columns and generate a PDF that still renders those columns with Aspose.Cells. | Convert an Excel file to PDF while keeping the hidden columns visible in the output by using Aspose.Cells HideColumns and SaveFormat.Pdf.
// Common Searches: Aspose.Cells C# hide specific columns before converting Excel to PDF | How to retain hidden columns in PDF output with Aspose.Cells .NET | C# example to hide columns K-M and export worksheet to PDF using Aspose.Cells | Export Excel to PDF including hidden columns using Aspose.Cells SaveFormat.Pdf
// Tags: Aspose.Cells HideColumns method usage | export worksheet to PDF with hidden columns | preserve column visibility in PDF conversion Aspose.Cells | C# Excel column hiding before PDF save | Aspose.Cells column visibility PDF output

using System;
using Aspose.Cells;

// Loads input.xlsx, hides columns K through M on the first worksheet using HideColumns, and saves the workbook as output.pdf where the hidden columns are still rendered in the PDF.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Hide columns K (index 10) through M (index 12) – total of 3 columns
        // The HideColumns method takes the zero‑based start index and the number of columns to hide
        workbook.Worksheets[0].Cells.HideColumns(10, 3);

        // Save the workbook as PDF. The PDF will include the hidden columns in the output.
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
