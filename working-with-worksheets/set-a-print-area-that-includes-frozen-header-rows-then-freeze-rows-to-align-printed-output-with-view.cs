// Title: How to set a print area that includes frozen header rows and repeat the header on each printed page with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to define a print area covering A1:D20, freeze the first worksheet row, and set that row as a repeated print title. | Show how to combine Worksheet.FreezePanes with PageSetup.PrintTitleRows so the printed pages display the same header that is frozen in the view. | Create an Excel workbook, populate sample data, configure the print area and frozen header, then save it as an .xlsx file using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set print area and freeze first row together | repeat header row on each printed page using Aspose.Cells .NET | align printed output with frozen panes in Aspose.Cells | C# Aspose.Cells example for print titles and freeze panes
// Tags: print area configuration Aspose.Cells C# | freeze panes implementation Aspose.Cells .NET | repeat header rows Aspose.Cells print titles | PageSetup.PrintTitleRows usage Aspose.Cells | save workbook as xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPrintAreaAndFreeze
{
    // // Creates a workbook, fills rows 1‑20 with sample data, sets the print area to A1:D20, freezes the first row, repeats that row as a print title on every page, and saves the file as PrintAreaAndFreeze.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data (including header rows)
                cells["A1"].PutValue("ID");
                cells["B1"].PutValue("Name");
                cells["C1"].PutValue("Department");
                cells["D1"].PutValue("Salary");

                for (int i = 2; i <= 20; i++)
                {
                    cells[$"A{i}"].PutValue(i - 1);
                    cells[$"B{i}"].PutValue($"Employee {i - 1}");
                    cells[$"C{i}"].PutValue("Sales");
                    cells[$"D{i}"].PutValue(3000 + (i - 2) * 100);
                }

                // Set the print area to include the header and data rows
                sheet.PageSetup.PrintArea = "A1:D20";

                // Freeze the first row (header) while scrolling
                // FreezePanes(row, column, totalRows, totalColumns)
                int totalRows = cells.MaxDataRow + 1;
                int totalColumns = cells.MaxDataColumn + 1;
                sheet.FreezePanes(1, 0, totalRows, totalColumns);

                // Repeat the header row on each printed page
                sheet.PageSetup.PrintTitleRows = "$1:$1";

                // Save the workbook
                string outputPath = "PrintAreaAndFreeze.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
