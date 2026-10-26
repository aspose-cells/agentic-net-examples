// Title: Create a macro‑free Excel report that lists all cells using the Hyperlink theme color with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells list cells that use the Hyperlink theme color without macros | how to generate a report of hyperlink‑colored cells in Excel using Aspose.Cells | iterate through workbook used range to find Hyperlink theme color cells in .NET | save an Excel file macro‑free after extracting cells with Hyperlink theme color using Aspose.Cells | retrieve A1 addresses of cells styled with Hyperlink theme color in C#
// Tags: detect hyperlink theme color Aspose.Cells | export cell addresses by theme color .NET | iterate used range cells Aspose.Cells | save workbook without macros Aspose.Cells | create Excel report programmatically C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The C# program loads a source workbook, scans every worksheet's used range for cells whose font theme color is Hyperlink, records each cell's A1 address, adds a new worksheet named "HyperlinkThemeColorReport", writes the collected addresses, auto‑fits the column, and saves the workbook as a macro‑free XLSX file using Aspose.Cells.
class HyperlinkThemeColorReport
{
    static void Main()
    {
        const string sourcePath = "SourceWorkbook.xlsx";
        const string reportPath = "HyperlinkThemeColorReport.xlsx";

        // Verify source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // List to hold addresses of cells using the Hyperlink theme color
            List<string> hyperlinkColorCells = new List<string>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in sourceWorkbook.Worksheets)
            {
                // Get the used range of the worksheet (fully qualified to avoid ambiguity)
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                // Iterate through each cell in the used range
                foreach (Cell cell in usedRange)
                {
                    // Get the cell's style
                    Style style = cell.GetStyle();

                    // Check if the font uses the Hyperlink theme color.
                    // Use string comparison to avoid reliance on a specific enum member.
                    if (style.Font.ThemeColor.ToString().Equals("Hyperlink", StringComparison.OrdinalIgnoreCase))
                    {
                        // Store the cell address in A1 style with sheet name
                        hyperlinkColorCells.Add($"{sheet.Name}!{cell.Name}");
                    }
                }
            }

            // Add a new worksheet for the report
            int reportIndex = sourceWorkbook.Worksheets.Add();
            Worksheet reportSheet = sourceWorkbook.Worksheets[reportIndex];
            reportSheet.Name = "HyperlinkThemeColorReport";

            // Write header
            reportSheet.Cells[0, 0].PutValue("Cells Using Hyperlink Theme Color");
            Style headerStyle = reportSheet.Cells[0, 0].GetStyle();
            headerStyle.Font.IsBold = true;
            reportSheet.Cells[0, 0].SetStyle(headerStyle);

            // Write the list of cell addresses starting from row 1 (second row)
            for (int i = 0; i < hyperlinkColorCells.Count; i++)
            {
                reportSheet.Cells[i + 1, 0].PutValue(hyperlinkColorCells[i]);
            }

            // Auto‑fit the column for better readability
            reportSheet.AutoFitColumn(0);

            // Save the workbook (macro‑free)
            sourceWorkbook.Save(reportPath, SaveFormat.Xlsx);
            Console.WriteLine($"Report saved to {reportPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
