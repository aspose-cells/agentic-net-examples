// Title: C# – Rename Active Worksheet to QuarterlyReport with Aspose.Cells (preserve cell data)
// Description: Demonstrates how to rename the currently active worksheet in an Aspose.Cells workbook to "QuarterlyReport" while keeping all existing cell values intact. The example creates a new workbook, accesses the active sheet via ActiveSheetIndex, writes sample data, sets the Worksheet.Name property, and saves the file as an XLSX document.
// Keywords: Aspose.Cells rename worksheet | C# active sheet name change | preserve cell data rename worksheet | Worksheet.Name property Aspose | save workbook after renaming sheet
// Common Searches: rename active worksheet Aspose.Cells C# | change worksheet name without losing data | how to set worksheet name in Aspose.Cells | preserve cell values when renaming sheet | Aspose.Cells rename sheet example
// Developer Intent: Rename the active worksheet to "QuarterlyReport" while ensuring no cell data is lost.
// Use Cases: Rename the default sheet after populating data before exporting a report. | Assign period‑specific names (e.g., Q1, Q2) to worksheets in automated financial statements. | Standardize worksheet names in multi‑sheet workbooks for better end‑user navigation.
// AI Prompts: Generate C# code that renames the active worksheet in an Aspose.Cells workbook to a custom name without affecting existing cell values. | Show how to use Worksheet.Name to change the active sheet's title and then save the workbook as XLSX. | Explain whether changing Worksheet.Name in Aspose.Cells impacts the data stored in cells.

using System;
using Aspose.Cells;

// Demonstrates how to rename the currently active worksheet in an Aspose.Cells workbook to "QuarterlyReport" while keeping all existing cell values intact. The example creates a new workbook, accesses the active sheet via ActiveSheetIndex, writes sample data, sets the Worksheet.Name property, and saves the file as an XLSX document.
class RenameActiveWorksheet
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();

        // Access the active worksheet using the active sheet index
        Worksheet activeSheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

        // Example data to show that cell contents are preserved after renaming
        activeSheet.Cells["A1"].PutValue("Sample Data");

        // Rename the active worksheet to "QuarterlyReport"
        activeSheet.Name = "QuarterlyReport";

        // Save the workbook (saving rule)
        workbook.Save("QuarterlyReport.xlsx", SaveFormat.Xlsx);
    }
}
