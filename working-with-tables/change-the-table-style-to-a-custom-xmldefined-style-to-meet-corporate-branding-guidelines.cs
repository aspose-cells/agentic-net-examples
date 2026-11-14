// Title: Assign a table style to an Excel ListObject using Aspose.Cells for .NET (C#)
// Common Searches: how to set TableStyleName for a ListObject with Aspose.Cells C# | c# change Excel table style programmatically using Aspose.Cells | apply corporate branding table style to an Excel worksheet via Aspose.Cells .NET | Aspose.Cells example for customizing table appearance in a workbook | assign built‑in or custom table style to Excel table in C# code
// Tags: TableStyleName assignment ListObject C# | customize Excel table style Aspose.Cells | apply branding table style .NET | modify ListObject appearance programmatically | Aspose.Cells table style API

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample loads or creates a workbook, ensures a ListObject exists on the first worksheet, sets its TableStyleName property to a chosen style (e.g., "TableStyleMedium2"), and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a minimal one if it does not exist.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet wsCreate = workbook.Worksheets[0];
                wsCreate.Name = "Sheet1";

                // Sample data for the table.
                wsCreate.Cells["A1"].PutValue("Header1");
                wsCreate.Cells["B1"].PutValue("Header2");
                wsCreate.Cells["C1"].PutValue("Header3");
                wsCreate.Cells["A2"].PutValue("Item1");
                wsCreate.Cells["B2"].PutValue(10);
                wsCreate.Cells["C2"].PutValue(DateTime.Today);

                workbook.Save(inputPath);
            }

            // Get the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve or create the first table (ListObject) on the sheet.
            ListObject table;
            if (worksheet.ListObjects.Count > 0)
            {
                table = worksheet.ListObjects[0];
            }
            else
            {
                // Create a new table spanning A1:C5 with a header row.
                int index = worksheet.ListObjects.Add(0, 0, 4, 2, true);
                table = worksheet.ListObjects[index];
            }

            // Apply a built‑in table style (replace with any desired style name).
            // Example: "TableStyleMedium2" is a standard style that exists in Excel.
            table.TableStyleName = "TableStyleMedium2";

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
