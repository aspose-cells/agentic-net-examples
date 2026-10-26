// Title: How to autofill a date series and apply a custom mm-dd-yyyy number format to a range using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells fill date series down column until row 10 | Aspose.Cells apply custom date format to a specific range in Excel | programmatically extend sequential dates in Excel with Aspose.Cells C# | set number format mm-dd-yyyy for cells A1:A10 using Aspose.Cells | create date series and style range in Aspose.Cells workbook C#
// Tags: autofill date series Aspose.Cells C# | custom number format range Aspose.Cells | date series generation Excel C# | apply mm-dd-yyyy style Aspose.Cells | range styling with Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, inserts three initial dates in column A, programmatically extends the series to A10 by adding one day to the previous date, defines a custom number format "mm-dd-yyyy", applies this format only to the range A1:A10, and saves the file as DateSeries.xlsx.
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

            // Initialize the first three dates in column A
            sheet.Cells["A1"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 2));
            sheet.Cells["A3"].PutValue(new DateTime(2023, 1, 3));

            // Fill the rest of the series up to A10
            for (int row = 4; row <= 10; row++)
            {
                // Get the previous date (row-2 because rows are 0‑based)
                DateTime previousDate = sheet.Cells[row - 2, 0].DateTimeValue;
                // Write the next date
                sheet.Cells[row - 1, 0].PutValue(previousDate.AddDays(1));
            }

            // Create a style with the desired date number format
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "mm-dd-yyyy";

            // Apply the style only to the number format of the filled range
            StyleFlag styleFlag = new StyleFlag { NumberFormat = true };
            Aspose.Cells.Range targetRange = sheet.Cells.CreateRange("A1:A10");
            targetRange.ApplyStyle(dateStyle, styleFlag);

            // Determine output path and ensure the directory exists
            string outputPath = "DateSeries.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
