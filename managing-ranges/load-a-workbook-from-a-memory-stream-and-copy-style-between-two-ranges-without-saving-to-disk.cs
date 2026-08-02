// Title: Copy cell style between ranges from a MemoryStream workbook – Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook in memory, apply a custom style to range A1:B2, save it to a MemoryStream, reload the workbook, add a new worksheet, and copy the style to range C3:D4 using Aspose.Cells without any disk I/O.
// Keywords: Aspose.Cells | CopyStyle | MemoryStream | C# | .NET | in‑memory Excel | range formatting | load workbook from stream | no disk I/O | style transfer between worksheets
// Common Searches: Aspose.Cells copy style from memory stream | load workbook from stream C# Aspose.Cells | copy cell formatting between ranges without saving file | CopyStyle method example Aspose.Cells .NET | in‑memory Excel style transfer Aspose
// Developer Intent: Load an Excel workbook from a MemoryStream and copy its cell style to another range without writing the file to disk.
// Use Cases: Apply a predefined style to a source range, keep the workbook in memory, and replicate the style on a different worksheet. | Build a web API that receives an Excel file as a byte array, modifies formatting in memory, and returns the updated file without temporary files. | Validate successful style copying by checking a specific attribute such as Font.IsBold on the destination cells.
// AI Prompts: Generate C# code that loads an Aspose.Cells workbook from a MemoryStream and copies the style from range A1:B2 to range C3:D4 on another worksheet. | Explain how to confirm that CopyStyle succeeded by inspecting a style property like Font.IsBold after the operation. | Provide an ASP.NET Core controller action that accepts an Excel file as a byte array, applies a custom style, copies it to a new sheet, and returns the modified workbook as a byte array.

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// Demonstrates how to create a workbook in memory, apply a custom style to range A1:B2, save it to a MemoryStream, reload the workbook, add a new worksheet, and copy the style to range C3:D4 using Aspose.Cells without any disk I/O.
class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create a workbook, apply a style, and save it to a memory stream ----------
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Define a style
            Style sampleStyle = sourceWorkbook.CreateStyle();
            sampleStyle.Font.Name = "Calibri";
            sampleStyle.Font.Size = 12;
            sampleStyle.Font.IsBold = true;
            sampleStyle.ForegroundColor = Color.LightBlue;
            sampleStyle.Pattern = BackgroundType.Solid;

            // Apply the style to a source range (A1:B2)
            Aspose.Cells.Range sourceRange = sourceSheet.Cells.CreateRange("A1:B2");
            sourceRange.SetStyle(sampleStyle);

            // Save the workbook into a MemoryStream (no disk I/O)
            using (MemoryStream memoryStream = sourceWorkbook.SaveToStream())
            {
                // Reset stream position for reading
                memoryStream.Position = 0;

                // ---------- Load the workbook from the memory stream ----------
                Workbook workbook = new Workbook(memoryStream);

                // Get the same source range from the loaded workbook
                Worksheet loadedSourceSheet = workbook.Worksheets[0];
                Aspose.Cells.Range loadedSourceRange = loadedSourceSheet.Cells.CreateRange("A1:B2");

                // Create a destination worksheet and range (C3:D4)
                int destSheetIndex = workbook.Worksheets.Add();
                Worksheet destSheet = workbook.Worksheets[destSheetIndex];
                destSheet.Name = "Destination";
                Aspose.Cells.Range destRange = destSheet.Cells.CreateRange("C3:D4");

                // Copy style from the source range to the destination range
                destRange.CopyStyle(loadedSourceRange);

                // Optional: verify that the style was copied
                Console.WriteLine("Destination range font bold: " +
                    destRange[0, 0].GetStyle().Font.IsBold);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
