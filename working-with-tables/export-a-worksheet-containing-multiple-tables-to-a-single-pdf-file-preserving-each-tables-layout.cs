// Title: Export multiple Excel tables on a single worksheet to one PDF file with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export worksheet containing several ListObjects to a single PDF | How to convert multiple Excel tables on the same sheet into one PDF using Aspose.Cells | Aspose.Cells save workbook as PDF while preserving table styles in C# | Export Excel tables to PDF with auto‑fit columns using Aspose.Cells .NET | Generate PDF from Excel sheet that has two ListObject tables in C#
// Tags: export worksheet tables to PDF Aspose.Cells | Aspose.Cells ListObject PDF conversion | preserve Excel table styles during PDF export | auto-fit columns for PDF layout Aspose.Cells | save workbook as single PDF with multiple tables

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a workbook, adds two ListObject tables with different styles on the same worksheet, auto‑fits the columns, ensures the output folder exists, and saves the entire sheet as a single PDF file named 'MultipleTables.pdf' using Aspose.Cells for .NET.
class ExportMultipleTablesToPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Populate data for the first table.
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(92);

            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Charlie");
            sheet.Cells["C4"].PutValue(78);

            sheet.Cells["A5"].PutValue(4);
            sheet.Cells["B5"].PutValue("Diana");
            sheet.Cells["C5"].PutValue(88);

            // Add the first table (range A1:C5) with headers.
            int table1Index = sheet.ListObjects.Add("Table1", "A1:C5", true);
            ListObject table1 = sheet.ListObjects[table1Index];
            table1.TableStyleType = TableStyleType.TableStyleMedium2;

            // Populate data for the second table.
            sheet.Cells["E1"].PutValue("Product");
            sheet.Cells["F1"].PutValue("Category");
            sheet.Cells["G1"].PutValue("Price");

            sheet.Cells["E2"].PutValue("Laptop");
            sheet.Cells["F2"].PutValue("Electronics");
            sheet.Cells["G2"].PutValue(1200);

            sheet.Cells["E3"].PutValue("Desk");
            sheet.Cells["F3"].PutValue("Furniture");
            sheet.Cells["G3"].PutValue(350);

            sheet.Cells["E4"].PutValue("Pen");
            sheet.Cells["F4"].PutValue("Stationery");
            sheet.Cells["G4"].PutValue(2.5);

            sheet.Cells["E5"].PutValue("Coffee");
            sheet.Cells["F5"].PutValue("Beverage");
            sheet.Cells["G5"].PutValue(5);

            // Add the second table (range E1:G5) with headers.
            int table2Index = sheet.ListObjects.Add("Table2", "E1:G5", true);
            ListObject table2 = sheet.ListObjects[table2Index];
            table2.TableStyleType = TableStyleType.TableStyleMedium9;

            // Adjust column widths for better PDF layout.
            sheet.AutoFitColumns();

            // Ensure the output directory exists.
            string outputPath = "MultipleTables.pdf";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the worksheet to PDF.
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF file saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
