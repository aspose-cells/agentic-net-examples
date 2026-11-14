// Title: How to ungroup columns and preserve original column widths in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# ungroup columns without changing column width | retain column widths after ungrouping columns in Excel with Aspose.Cells | save original column widths before ungrouping columns using Aspose.Cells .NET | C# code to ungroup previously grouped columns and keep column sizes in Aspose.Cells | Aspose.Cells set column width after ungrouping columns in workbook
// Tags: ungroup columns Aspose.Cells .NET | preserve column width Aspose.Cells | column width restoration after ungrouping | Excel workbook column width consistency Aspose.Cells | grouped columns hide ungroup Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, groups columns B‑D, modifies data inside the group, records the widths of all used columns, then ungroups the columns and reapplies the saved widths so the worksheet layout stays unchanged before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["C1"].PutValue("Header3");
            sheet.Cells["D1"].PutValue("Header4");

            // Group columns B to D (indexes 1 to 3) and hide them
            sheet.Cells.GroupColumns(1, 3, true);

            // Modify data within the grouped columns
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue(20);
            sheet.Cells["D2"].PutValue(30);

            // Store current column widths for all used columns
            int maxColumn = sheet.Cells.MaxColumn + 1;
            double[] originalWidths = new double[maxColumn];
            for (int i = 0; i < maxColumn; i++)
            {
                originalWidths[i] = sheet.Cells.GetColumnWidth(i);
            }

            // Ungroup the previously grouped columns
            sheet.Cells.UngroupColumns(1, 3);

            // Reapply the original column widths to keep them consistent
            for (int i = 0; i < maxColumn; i++)
            {
                sheet.Cells.SetColumnWidth(i, originalWidths[i]);
            }

            // Define output file path
            string outputPath = "Output.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
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
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
