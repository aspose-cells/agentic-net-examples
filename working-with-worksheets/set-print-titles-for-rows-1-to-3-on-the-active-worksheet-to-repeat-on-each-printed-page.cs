// Title: How to set rows 1‑3 as repeating print titles on an Aspose.Cells worksheet in C#
// AI Prompts: Write C# code that uses Aspose.Cells to configure the first three rows as print titles so they appear on every printed page. | Demonstrate how to assign PageSetup.PrintTitleRows for rows 1 through 3 and save the workbook as an .xlsx file with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# repeat first three rows on each printed page | Set print title rows 1 to 3 in an Excel workbook using Aspose.Cells .NET | How to use PageSetup.PrintTitleRows to create repeating header rows in Aspose.Cells | C# example for printing repeated header rows with Aspose.Cells | Configure worksheet print titles across pages in Aspose.Cells for .NET
// Tags: Aspose.Cells PageSetup.PrintTitleRows property | C# Aspose.Cells configure print title rows | Excel workbook repeat rows on print Aspose.Cells | Aspose.Cells worksheet print setup .NET | save workbook with print titles Aspose.Cells

using Aspose.Cells;
using System;

// The program creates a new workbook, accesses the first worksheet, configures rows 1‑3 as print titles using the PageSetup.PrintTitleRows property so they repeat on every printed page, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Get the active worksheet (first worksheet by default)
        Worksheet worksheet = workbook.Worksheets[0];

        // Set rows 1 to 3 as print titles so they repeat on each printed page
        // The format is "$StartRow:$EndRow"
        worksheet.PageSetup.PrintTitleRows = "$1:$3";

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
