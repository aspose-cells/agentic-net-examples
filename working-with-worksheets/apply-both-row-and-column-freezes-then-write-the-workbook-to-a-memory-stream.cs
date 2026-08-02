// Title: Freeze rows and columns and save an Aspose.Cells workbook to a MemoryStream (C#)
// Description: Demonstrates how to create a new Workbook, apply FreezePanes to lock the first two rows and three columns at cell D3, save the workbook to a MemoryStream with SaveToStream, reset the stream position, and optionally write the stream to a .xlsx file.
// Keywords: Aspose.Cells FreezePanes C# | SaveToStream Aspose.Cells | Excel MemoryStream .NET | freeze rows columns Aspose | export workbook to stream | write stream to file C# | Aspose.Cells example
// Common Searches: Aspose.Cells freeze first rows and columns | How to use FreezePanes in Aspose.Cells C# | Save Aspose.Cells workbook to MemoryStream | Convert Aspose.Cells workbook to byte array | Write Aspose.Cells stream to a file | C# example for freezing panes in Excel
// Developer Intent: Create an Excel workbook with specific rows and columns frozen and obtain the result as a MemoryStream for further processing or delivery.
// Use Cases: Generate a report template with frozen header rows, keep it in memory, and return it as a downloadable file in a web API. | Produce a large spreadsheet on a server, apply FreezePanes, store it in a MemoryStream, and upload directly to cloud storage without a temporary file. | Create an Excel workbook in memory, freeze panes for better navigation, and send the MemoryStream as an email attachment.
// AI Prompts: Provide C# code using Aspose.Cells to freeze the first two rows and three columns and return the workbook as a MemoryStream. | Show how to save an Aspose.Cells workbook to a MemoryStream, reset its position, and then write it to a .xlsx file on disk. | Explain the parameters of FreezePanes in Aspose.Cells and how to adjust them to freeze rows above and columns to the left of a given cell.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a new Workbook, apply FreezePanes to lock the first two rows and three columns at cell D3, save the workbook to a MemoryStream with SaveToStream, reset the stream position, and optionally write the stream to a .xlsx file.
public class FreezeAndSaveDemo
{
    // Creates a workbook, applies row and column freezes, and returns it as a MemoryStream.
    public static MemoryStream CreateWorkbookWithFreeze()
    {
        try
        {
            // Initialize a new workbook (default format is Xlsx).
            Workbook workbook = new Workbook();

            // Get the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Freeze panes: freeze the first 2 rows and first 3 columns at cell D3.
            worksheet.FreezePanes(2, 3, 2, 3);

            // Save the workbook to a memory stream.
            MemoryStream stream = workbook.SaveToStream();

            // Ensure the stream position is reset for reading.
            stream.Position = 0;
            return stream;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating workbook: {ex.Message}");
            return null;
        }
    }

    // Entry point for the console application.
    public static void Main()
    {
        try
        {
            MemoryStream workbookStream = CreateWorkbookWithFreeze();
            if (workbookStream == null)
            {
                Console.WriteLine("Workbook creation failed.");
                return;
            }

            string outputPath = "FreezeDemo.xlsx";

            // Write the memory stream to a file.
            using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                workbookStream.CopyTo(file);
            }

            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
