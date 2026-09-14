// Title: Apply a built‑in TableStyleMedium2 to an Aspose.Cells ListObject while preserving existing cell formatting in C#
// AI Prompts: Generate C# code that creates a worksheet, fills cells A1:B3 with sample data, converts the range into a ListObject, and sets its TableStyleType to TableStyleMedium2 using Aspose.Cells. | Write a C# snippet that applies a predefined Excel table style to an Aspose.Cells ListObject without altering the original cell formats. | Provide C# code to save the styled workbook as an .xlsx file, ensuring the output directory is created if it does not already exist.
// Common Searches: how to set TableStyleMedium2 for an Excel table created with Aspose.Cells in C# | preserving cell formatting when applying a table style with Aspose.Cells | Aspose.Cells C# add ListObject and apply built‑in table style | save Aspose.Cells workbook to a specific folder and create folder if missing | apply predefined table style to ListObject without losing header formatting Aspose
// Tags: apply TableStyleMedium2 to ListObject Aspose.Cells | preserve cell formatting when styling Excel table C# | create ListObject from range Aspose.Cells | save workbook to .xlsx with directory creation Aspose.Cells | use TableStyleType enum Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates creating a workbook, adding sample data, converting range A1:B3 into a ListObject, applying the built‑in TableStyleMedium2 style while keeping original cell formatting, and saving the file with automatic output folder creation.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(25);

            // Define the range that will become a table (A1:B3)
            var tableRange = sheet.Cells.CreateRange("A1:B3");

            // Add a ListObject (Excel table) based on the defined range, indicating that the first row has headers
            int tableIndex = sheet.ListObjects.Add(
                tableRange.FirstRow,
                tableRange.FirstColumn,
                tableRange.RowCount,
                tableRange.ColumnCount,
                true);

            var table = sheet.ListObjects[tableIndex];

            // Apply a predefined table style (e.g., TableStyleMedium2)
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Save the workbook to a file
            string outputFile = "StyledTable.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputFile)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
