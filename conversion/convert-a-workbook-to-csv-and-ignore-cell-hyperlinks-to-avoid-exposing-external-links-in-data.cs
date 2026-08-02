// Title: Convert Excel to CSV without Hyperlinks using Aspose.Cells for .NET
// Description: Loads an Excel workbook, removes all worksheet hyperlinks and external link references, then saves the first worksheet as a CSV file with Aspose.Cells for C#. This prevents URLs from appearing in the exported data.
// Keywords: Aspose.Cells CSV export | remove hyperlinks C# | clear external links Aspose | Excel to CSV without URLs | Aspose.Cells workbook conversion
// Common Searches: Aspose.Cells export to CSV without hyperlinks | C# remove Excel hyperlinks before CSV conversion | how to strip external links from workbook using Aspose | save first worksheet as CSV Aspose.Cells | CSV export ignoring Excel hyperlinks
// Developer Intent: Generate a CSV file from an Excel workbook while ensuring that no hyperlink or external link data is included.
// Use Cases: Produce clean CSV datasets for analytics pipelines that must not contain URLs. | Create CSV reports from Excel templates that embed links, avoiding accidental exposure of external resources. | Prepare data for systems that reject hyperlink fields by stripping them prior to CSV conversion.
// AI Prompts: Write C# code with Aspose.Cells to export multiple selected worksheets to separate CSV files after removing all hyperlinks and external links. | Show how to keep numeric formatting and cell values intact while clearing hyperlinks before saving an Excel workbook as CSV. | Provide a snippet that logs the count of hyperlinks removed from each worksheet during CSV export.

using System;
using Aspose.Cells;

// Loads an Excel workbook, removes all worksheet hyperlinks and external link references, then saves the first worksheet as a CSV file with Aspose.Cells for C#. This prevents URLs from appearing in the exported data.
class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Remove all hyperlinks from each worksheet to prevent external links from being exported
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.Hyperlinks.Clear();
        }

        // Additionally clear any external link references that might exist
        workbook.Worksheets.ExternalLinks.Clear(true);

        // Save the workbook as CSV; only the first worksheet is saved in CSV format
        workbook.Save("output.csv", SaveFormat.Csv);
    }
}
