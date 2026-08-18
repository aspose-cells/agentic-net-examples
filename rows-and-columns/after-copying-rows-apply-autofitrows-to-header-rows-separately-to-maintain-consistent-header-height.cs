// Title: Copy rows and auto‑fit header rows separately with Aspose.Cells for .NET
// Description: Shows how to copy a range of data rows, apply Worksheet.AutoFitRows only to header rows (rows 0‑1) to keep their height consistent, then auto‑fit the remaining rows up to the last used row, and finally save the workbook.
// Keywords: Aspose.Cells copy rows C# | Worksheet.AutoFitRows header | preserve header height Aspose.Cells | copy rows and autofit rows .NET | Excel row copy Aspose.Cells | C# Aspose.Cells AutoFitRows range | Aspose.Cells copy rows example | auto fit specific rows Aspose.Cells
// Common Searches: Aspose.Cells copy rows example | AutoFitRows specific rows C# | keep header height after copying rows Aspose.Cells | copy rows and auto fit rows Aspose.Cells .NET | Worksheet.AutoFitRows range usage | C# copy rows in Excel with Aspose.Cells
// Developer Intent: Copy a block of rows and then auto‑fit only the header rows while allowing all other rows to adjust automatically.
// Use Cases: Duplicate a table within the same worksheet and keep the original header rows at a fixed height. | Generate a report where header rows must remain uniform after copying data rows to a new location. | Create a template that copies data rows for multiple sections while preserving consistent header formatting.
// AI Prompts: Provide C# code that copies rows 2‑5 to rows 6‑9 and auto‑fits only rows 0‑1 using Aspose.Cells. | Show an example of applying Worksheet.AutoFitRows to separate header and data ranges after copying rows in a workbook. | Explain how to preserve header row height while auto‑fitting all other rows after using Cells.CopyRows in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to copy a range of data rows, apply Worksheet.AutoFitRows only to header rows (rows 0‑1) to keep their height consistent, then auto‑fit the remaining rows up to the last used row, and finally save the workbook.
    public class CopyRowsAndAutoFitHeaderDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (source and destination are the same for simplicity)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // 1. Populate sample data
            // -------------------------------------------------
            // Header rows (rows 0 and 1)
            cells["A1"].PutValue("Header Column 1");
            cells["B1"].PutValue("Header Column 2");
            cells["A2"].PutValue("Sub Header 1");
            cells["B2"].PutValue("Sub Header 2");

            // Data rows (rows 2 to 5)
            for (int i = 2; i <= 5; i++)
            {
                cells[i, 0].PutValue($"Data Row {i - 1} - Column A with a relatively long text to test autofit");
                cells[i, 1].PutValue($"Data Row {i - 1} - Column B");
            }

            // -------------------------------------------------
            // 2. Copy rows 2-5 (data rows) to rows 6-9 (below the original data)
            // -------------------------------------------------
            // Parameters: sourceCells, sourceRowIndex, destinationRowIndex, rowNumber
            cells.CopyRows(cells, 2, 6, 4); // copies 4 rows starting from row index 2 to row index 6

            // -------------------------------------------------
            // 3. AutoFit only the header rows to keep their height consistent
            // -------------------------------------------------
            // Header rows are 0 and 1 (inclusive)
            sheet.AutoFitRows(0, 1);

            // -------------------------------------------------
            // 4. AutoFit the rest of the rows (data rows and copied rows)
            // -------------------------------------------------
            // Data rows start at row 2 and go through the last used row
            int lastRow = cells.MaxDataRow; // gets the index of the last row that contains data
            sheet.AutoFitRows(2, lastRow);

            // -------------------------------------------------
            // 5. Save the workbook
            // -------------------------------------------------
            string outputPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "CopyRowsAutoFitHeaderDemo.xlsx");
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}
