// Title: How to group rows and auto‑fit their height in an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells group rows and auto fit row height example | programmatically hide and auto‑fit grouped rows in Excel with Aspose.Cells | Aspose.Cells .NET group a range of rows and adjust row height | auto fit rows after grouping using Aspose.Cells C# code sample | group rows 3 to 7 and auto‑fit in Aspose.Cells workbook
// Tags: Aspose.Cells row grouping | Aspose.Cells auto fit rows | Excel hide grouped rows C# | Aspose.Cells adjust row height | group rows with auto fit .NET

using Aspose.Cells;
using System;
using System.IO;

// Creates a new workbook, fills column A with sample data, groups rows 3‑7 (zero‑based indexes 2‑6) while hiding them, then auto‑fits the height of those grouped rows, and saves the result as GroupedRows.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Fill sample data into rows 0‑9, column A
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue($"Row {i + 1}");
            }

            // Define the rows to group (e.g., rows 3‑7 in Excel, indexes 2‑6 in zero‑based)
            int startRow = 2;                     // zero‑based start index
            int endRow = 6;                       // zero‑based end index
            int totalRows = endRow - startRow + 1;

            // Group the rows and hide them (true = hidden)
            sheet.Cells.GroupRows(startRow, totalRows, true);

            // Auto‑fit the grouped rows to adjust their height
            sheet.AutoFitRows(startRow, totalRows);

            // Ensure the output directory exists
            string outputPath = "GroupedRows.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
