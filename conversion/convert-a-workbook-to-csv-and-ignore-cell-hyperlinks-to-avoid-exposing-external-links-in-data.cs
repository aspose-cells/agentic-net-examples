// Title: Export Excel to CSV without Hyperlinks using Aspose.Cells for .NET
// Description: Load an Excel workbook, clear external links and worksheet hyperlink collections, then save it as a CSV file so the output contains no hyperlink data.
// Keywords: Aspose.Cells CSV export | remove hyperlinks Aspose.Cells | clear external links .NET | Excel to CSV without links | Aspose.Cells hyperlink removal
// Common Searches: Aspose.Cells export to CSV without hyperlinks | C# remove Excel hyperlinks before CSV conversion | How to clear external links in Aspose.Cells | Save workbook as CSV ignoring hyperlinks .NET | Strip hyperlinks from Excel when converting to CSV
// Developer Intent: Generate a CSV file from an Excel workbook while ensuring that no hyperlink information is included in the exported data.
// Use Cases: Sanitizing data for downstream analytics by removing clickable URLs. | Creating compliance‑ready CSV reports that must not expose external links. | Automating batch conversion of multiple workbooks to clean CSV files.
// AI Prompts: Write C# code with Aspose.Cells to load an .xlsx, clear all external links and worksheet hyperlinks, and save as CSV. | Explain the impact of Workbook.Worksheets.ExternalLinks.Clear(true) and Worksheet.Hyperlinks.Clear() on CSV output. | Provide a step‑by‑step guide to batch‑process a folder of Excel files, stripping hyperlinks and exporting each to CSV using Aspose.Cells.

using System;
using Aspose.Cells;

// Load an Excel workbook, clear external links and worksheet hyperlink collections, then save it as a CSV file so the output contains no hyperlink data.
class WorkbookToCsvWithoutHyperlinks
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(sourcePath);

        // Remove all external links (including hyperlinks) from the workbook
        // The boolean parameter updates references to local ones when possible
        workbook.Worksheets.ExternalLinks.Clear(true);

        // Additionally clear any hyperlink collections on each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.Hyperlinks.Clear();
        }

        // Save the workbook as CSV; hyperlinks will not be present in the output
        string csvPath = "output.csv";
        workbook.Save(csvPath, SaveFormat.Csv);
    }
}
