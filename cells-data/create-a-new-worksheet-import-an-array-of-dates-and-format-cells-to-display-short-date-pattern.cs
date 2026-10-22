// Title: Add a new worksheet, import a DateTime array into column A, and format cells with the short date style using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook, adds a worksheet named "Dates", inserts a DateTime[] into column A, and applies the built‑in short date number format (14) to those cells. | Show how to set the short date number format for a range of date cells in an Aspose.Cells workbook and then save the workbook as an .xlsx file.
// Common Searches: aspnet how to write a DateTime array to an Aspose.Cells worksheet and display short dates | c# Aspose.Cells set built‑in short date format for column of dates | example of adding a new sheet and formatting dates with number format 14 in Aspose.Cells | save workbook with formatted date cells using Aspose.Cells for .NET
// Tags: add worksheet and import DateTime array Aspose.Cells | apply built‑in short date number format Aspose.Cells | format date cells with number 14 Aspose.Cells | save workbook as xlsx Aspose.Cells | populate column A with dates Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Creates a new workbook, adds a worksheet named "Dates", writes a predefined DateTime array into column A, applies the built‑in short date number format (14) to each cell, and saves the file as DatesWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Initialize a new workbook
            Workbook workbook = new Workbook();

            // Add a new worksheet named "Dates"
            Worksheet sheet = workbook.Worksheets.Add("Dates");

            // Prepare an array of DateTime values
            DateTime[] dates = new DateTime[]
            {
                new DateTime(2023, 1, 1),
                new DateTime(2023, 2, 14),
                new DateTime(2023, 3, 31),
                new DateTime(2023, 4, 15),
                new DateTime(2023, 5, 20)
            };

            // Insert dates into column A (zero‑based indices)
            for (int i = 0; i < dates.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(dates[i]);
            }

            // Create a style for short date pattern (built‑in format 14)
            Style shortDateStyle = workbook.CreateStyle();
            shortDateStyle.Number = 14;

            // Apply the style to each cell containing a date
            for (int i = 0; i < dates.Length; i++)
            {
                sheet.Cells[i, 0].SetStyle(shortDateStyle);
            }

            // Save the workbook to a file
            string outputPath = "DatesWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
