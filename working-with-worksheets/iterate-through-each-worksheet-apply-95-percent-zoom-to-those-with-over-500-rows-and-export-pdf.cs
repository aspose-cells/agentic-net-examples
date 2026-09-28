// Title: Apply 95% zoom to worksheets that contain more than 500 rows and export the workbook as PDF with Aspose.Cells for .NET
// AI Prompts: Loop through all worksheets in a workbook, check if the data rows exceed 500, set each sheet's Zoom property to 95, and then save the workbook as a PDF using Aspose.Cells in C#. | Using Aspose.Cells, adjust the zoom level of each sheet based on a row‑count threshold and generate a PDF output in a single operation.
// Common Searches: Aspose.Cells C# set worksheet zoom when row count is greater than 500 | How to export an Excel file to PDF after changing zoom for large sheets with Aspose.Cells | Iterate over worksheets and apply conditional zoom before PDF conversion in .NET | C# code to adjust zoom level of sheets with more than 500 rows using Aspose.Cells | Save workbook as PDF with custom zoom settings per worksheet Aspose.Cells
// Tags: worksheet zoom based on row count Aspose.Cells | conditional zoom setting C# | Aspose.Cells PDF export with custom zoom | iterate worksheets set zoom threshold | Aspose.Cells Zoom property usage

using Aspose.Cells;
using System;

// // Loads an Excel file, sets a 95% zoom on any worksheet containing more than 500 rows, and saves the workbook as a PDF using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the number of rows that contain data (zero‑based index + 1)
            int rowCount = sheet.Cells.MaxDataRow + 1;

            // If the worksheet has more than 500 rows, set zoom to 95%
            if (rowCount > 500)
            {
                sheet.Zoom = 95;
            }
        }

        // Export the workbook to a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
