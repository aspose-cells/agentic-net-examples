// Title: Hide blank worksheets in multiple Excel workbooks and overwrite the files using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a collection of .xlsx file paths, iterates each workbook's worksheets, hides any sheet whose MaxDataRow and MaxDataColumn indicate no data, and saves the workbook back to the same location. | Create a reusable method that accepts a workbook, checks every worksheet for content using MaxDataRow/MaxDataColumn, toggles the IsVisible property accordingly, and overwrites the original Excel file.
// Common Searches: Aspose.Cells C# hide worksheets with no data in a batch of Excel files | how to programmatically hide empty sheets and save the workbook using Aspose.Cells | check worksheet emptiness with MaxDataRow and MaxDataColumn Aspose.Cells | overwrite original .xlsx after changing sheet visibility Aspose.Cells | process multiple workbooks and hide blank worksheets using Aspose.Cells for .NET
// Tags: hide empty worksheets Aspose.Cells | batch process Excel workbooks C# | detect worksheet data presence MaxDataRow Aspose.Cells | set worksheet visibility programmatically Aspose.Cells | overwrite workbook after sheet visibility change Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;

// The program loads each Excel file from a predefined list, examines every worksheet with Aspose.Cells, determines if the sheet contains any data by checking MaxDataRow and MaxDataColumn, hides sheets that are empty (or ensures they remain visible when data exists), and then saves the workbook back to the original .xlsx file, effectively overwriting it.
class Program
{
    static void Main()
    {
        // Collection of workbook file paths to process
        var workbookFiles = new List<string>
        {
            "Workbook1.xlsx",
            "Workbook2.xlsx"
            // Add additional workbook paths as needed
        };

        foreach (var filePath in workbookFiles)
        {
            // Load the workbook from file
            var workbook = new Workbook(filePath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Check if the worksheet contains any data
                bool hasData = sheet.Cells.MaxDataRow >= 0 && sheet.Cells.MaxDataColumn >= 0;

                // Hide the worksheet if it has no data; otherwise ensure it is visible
                sheet.IsVisible = hasData;
            }

            // Save the workbook back to the original file (overwrite)
            workbook.Save(filePath, SaveFormat.Xlsx);
        }
    }
}
