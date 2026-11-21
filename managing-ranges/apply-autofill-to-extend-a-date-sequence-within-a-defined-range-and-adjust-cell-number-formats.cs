// Title: Autofill a date series from A1 to A10 and apply built‑in date format 14 using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells autofill date series to extend range A1:A10 | How to apply built‑in date number format 14 to a range with Aspose.Cells | Using AutoFillType.Default to fill dates in an Excel worksheet via Aspose.Cells | Set number format for a filled range with StyleFlag in Aspose.Cells C# | Create and style a date column programmatically with Aspose.Cells .NET
// Tags: extend range with AutoFill Aspose.Cells | built‑in date format 14 styling Aspose.Cells | StyleFlag numberformat configuration Aspose.Cells | programmatic date column creation Aspose.Cells | Excel workbook save using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example creates a workbook, writes a three‑day date series to cells A1‑A3, uses Aspose.Cells AutoFill to extend the series to A10, defines a style with built‑in number format 14, applies the style to the filled range via a StyleFlag, and saves the file as 'AutofillDateSequence.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Initialize a date sequence in cells A1:A3
            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 3; i++)
            {
                sheet.Cells[i, 0].PutValue(startDate.AddDays(i));
            }

            // Extend the date sequence to A10 using Autofill
            AsposeRange source = sheet.Cells.CreateRange("A1:A3");
            AsposeRange destination = sheet.Cells.CreateRange("A1:A10");
            source.AutoFill(destination, AutoFillType.Default);

            // Define a date number format (built‑in format 14 = "m/d/yyyy")
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // or dateStyle.Custom = "mm-dd-yyyy";

            // Apply the date format to the filled range
            AsposeRange filledRange = sheet.Cells.CreateRange("A1:A10");
            StyleFlag flag = new StyleFlag { NumberFormat = true };
            filledRange.ApplyStyle(dateStyle, flag);

            // Determine output file path and ensure the directory exists
            string outputPath = "AutofillDateSequence.xlsx";
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
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
