using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// Alias to avoid conflict with System.Range (C# 8+)
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        try
        {
            // ------------------------------------------------------------
            // 1. Create a workbook in memory and apply a style to a source range
            // ------------------------------------------------------------
            using (Workbook sourceWorkbook = new Workbook())
            {
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

                // Define source range A1:B2
                AsposeRange sourceRange = sourceSheet.Cells.CreateRange("A1:B2");

                // Create and configure a style
                Style style = sourceWorkbook.CreateStyle();
                style.Font.Name = "Calibri";
                style.Font.Size = 12;
                style.Font.IsBold = true;
                style.ForegroundColor = Color.LightBlue;
                style.Pattern = BackgroundType.Solid;

                // Apply the style to the source range
                sourceRange.SetStyle(style);

                // ------------------------------------------------------------
                // 2. Save the workbook to a MemoryStream (no disk I/O)
                // ------------------------------------------------------------
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    sourceWorkbook.Save(memoryStream, SaveFormat.Xlsx);
                    memoryStream.Position = 0; // reset for reading

                    // ------------------------------------------------------------
                    // 3. Load a workbook from the MemoryStream
                    // ------------------------------------------------------------
                    using (Workbook targetWorkbook = new Workbook(memoryStream))
                    {
                        Worksheet targetSheet = targetWorkbook.Worksheets[0];

                        // ------------------------------------------------------------
                        // 4. Define a destination range and copy the style from source range
                        // ------------------------------------------------------------
                        AsposeRange destinationRange = targetSheet.Cells.CreateRange("C3:D4");
                        destinationRange.CopyStyle(sourceRange);

                        // ------------------------------------------------------------
                        // 5. (Optional) Verify that the style was copied
                        // ------------------------------------------------------------
                        Style copiedStyle = targetSheet.Cells["C3"].GetStyle();
                        Console.WriteLine($"Copied Font Bold: {copiedStyle.Font.IsBold}");
                        Console.WriteLine($"Copied Foreground Color: {copiedStyle.ForegroundColor}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Runtime safety: report any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}