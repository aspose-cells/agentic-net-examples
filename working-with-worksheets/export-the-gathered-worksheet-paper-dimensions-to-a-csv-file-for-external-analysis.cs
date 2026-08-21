// Title: Export Worksheet Paper Size to CSV with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, reads each worksheet's PageSetup PaperWidth and PaperHeight (in inches), writes the sheet name and dimensions to a CSV file, and saves it for external analysis.
// Keywords: Aspose.Cells | C# | .NET | export paper size | worksheet page setup | CSV export | PaperWidth | PaperHeight | Excel to CSV | print layout data
// Common Searches: Aspose.Cells get worksheet paper size C# | Export Excel page setup dimensions to CSV | How to read PaperWidth PaperHeight with Aspose.Cells | Save worksheet print size as CSV .NET | Extract worksheet print layout using Aspose.Cells
// Developer Intent: Extract each worksheet's print paper dimensions and output them to a CSV file.
// Use Cases: Audit print layout across all sheets in a workbook | Feed paper size data to a reporting system | Validate page setup before batch printing | Create an inventory of worksheet dimensions for migration projects
// AI Prompts: Generate C# code using Aspose.Cells that lists every worksheet's PaperWidth, PaperHeight, and Orientation in a CSV file. | Modify the example to include margins and a custom CSV delimiter such as a semicolon. | Explain how to convert dimensions from inches to millimeters when exporting worksheet paper sizes with Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an Excel workbook, reads each worksheet's PageSetup PaperWidth and PaperHeight (in inches), writes the sheet name and dimensions to a CSV file, and saves it for external analysis.
class ExportPaperDimensions
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        string sourcePath = "input.xlsx";
        Workbook sourceWorkbook = new Workbook(sourcePath);   // load rule

        // Create a new workbook that will hold the CSV data
        Workbook csvWorkbook = new Workbook();                // create rule
        Worksheet sheet = csvWorkbook.Worksheets[0];

        // Write CSV header
        sheet.Cells["A1"].PutValue("Worksheet");
        sheet.Cells["B1"].PutValue("PaperWidthInches");
        sheet.Cells["C1"].PutValue("PaperHeightInches");

        int rowIndex = 1; // zero‑based index for the next data row

        // Iterate through each worksheet in the source workbook
        foreach (Worksheet ws in sourceWorkbook.Worksheets)
        {
            // Retrieve paper dimensions (in inches) from the worksheet's PageSetup
            double width = ws.PageSetup.PaperWidth;
            double height = ws.PageSetup.PaperHeight;

            // Populate the CSV sheet with the gathered information
            sheet.Cells[rowIndex, 0].PutValue(ws.Name);
            sheet.Cells[rowIndex, 1].PutValue(width);
            sheet.Cells[rowIndex, 2].PutValue(height);
            rowIndex++;
        }

        // Save the populated workbook as a CSV file (uses the save rule)
        csvWorkbook.Save("paper_dimensions.csv", SaveFormat.Csv);
    }
}
