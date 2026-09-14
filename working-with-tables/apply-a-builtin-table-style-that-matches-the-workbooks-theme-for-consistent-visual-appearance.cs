// Title: Apply a built‑in themed table style to a ListObject in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Create a ListObject from a specified cell range and set its TableStyleType to a built‑in style that inherits the workbook's theme with Aspose.Cells in C#. | Load or generate an .xlsx file, add a table covering A1:D5, assign TableStyleMedium2, and save the workbook using Aspose.Cells for .NET.
// Common Searches: aspnet cells apply built‑in table style that matches workbook theme | c# set ListObject TableStyleType to TableStyleMedium2 using Aspose.Cells | how to add a styled table to an Excel file with Aspose.Cells .NET | apply theme‑aware table formatting in Aspose.Cells C# example | Aspose.Cells create table from range and use workbook theme style
// Tags: Aspose.Cells ListObject built‑in table style | C# TableStyleMedium2 Aspose.Cells | Excel workbook theme table formatting Aspose.Cells | create table from range Aspose.Cells .NET | apply themed table style to Excel worksheet C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The program loads (or creates) an Excel workbook, adds a ListObject covering cells A1:D5, sets its TableStyleType to TableStyleMedium2—a built‑in style that follows the workbook's theme—and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                var tempSheet = tempWb.Worksheets[0];

                // Populate sample data (A1:D5)
                for (int r = 0; r < 5; r++)
                {
                    for (int c = 0; c < 4; c++)
                    {
                        tempSheet.Cells[r, c].PutValue($"R{r + 1}C{c + 1}");
                    }
                }

                tempWb.Save(inputPath);
            }

            // Load the existing workbook (lifecycle rule: load)
            var workbook = new Workbook(inputPath);

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Define the range that will become a table (A1:D5)
            int firstRow = 0;      // Row index for A1
            int firstColumn = 0;   // Column index for A1
            int totalRows = 5;     // Number of rows (A1 to A5)
            int totalColumns = 4;  // Number of columns (A to D)

            // Add a ListObject (table) to the worksheet (hasHeaders = true)
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            var table = sheet.ListObjects[tableIndex];

            // Apply a built‑in table style that follows the workbook's theme
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Save the modified workbook (lifecycle rule: save)
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
