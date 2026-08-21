// Title: Batch Hide Empty Worksheets in Multiple Excel Workbooks with Aspose.Cells for .NET
// Description: Loads each workbook from a supplied list, scans every worksheet, marks sheets whose MaxDataRow and MaxDataColumn are -1 as invisible, and overwrites the original files using Aspose.Cells.
// Keywords: Aspose.Cells | C# Excel automation | hide empty worksheets | batch process workbooks | MaxDataRow | MaxDataColumn | worksheet visibility | save workbook .NET | multiple Excel files
// Common Searches: Aspose.Cells hide blank sheets in C# | process several Excel files and hide empty tabs | detect empty worksheet using MaxDataRow Aspose | overwrite original workbook after modifications Aspose.Cells | automate Excel cleanup with .NET
// Developer Intent: Automatically conceal all worksheets that contain no data in each workbook and persist the changes to the same files.
// Use Cases: Prepare distribution‑ready reports by removing unused tabs across a batch of workbooks. | Reduce file size and improve navigation in template files before they are shared with end users. | Clean up user‑generated Excel exports that may contain placeholder sheets left empty.
// AI Prompts: Create a C# method that accepts a collection of Excel paths, hides empty worksheets using Aspose.Cells, and returns a count of hidden sheets per file. | Write code that logs the filenames and sheet names that were hidden while processing multiple workbooks with Aspose.Cells. | Show an alternative way to identify empty worksheets in Aspose.Cells by using the UsedRange property instead of MaxDataRow/MaxDataColumn.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace HideEmptyWorksheetsDemo
{
    // Loads each workbook from a supplied list, scans every worksheet, marks sheets whose MaxDataRow and MaxDataColumn are -1 as invisible, and overwrites the original files using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // List of workbook file paths to process
            List<string> workbookFiles = new List<string>
            {
                "Book1.xlsx",
                "Book2.xlsx",
                // Add more file paths as needed
            };

            foreach (string filePath in workbookFiles)
            {
                // Load the workbook (create/load rule)
                Workbook workbook = new Workbook(filePath);

                // Iterate over all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Determine if the worksheet contains any data
                    // MaxDataRow and MaxDataColumn return -1 when there is no data
                    bool hasNoData = sheet.Cells.MaxDataRow == -1 && sheet.Cells.MaxDataColumn == -1;

                    if (hasNoData)
                    {
                        // Hide the worksheet (set visibility to false)
                        sheet.IsVisible = false;
                    }
                }

                // Save the modified workbook back to the same file (save rule)
                workbook.Save(filePath, SaveFormat.Xlsx);
            }
        }
    }
}
