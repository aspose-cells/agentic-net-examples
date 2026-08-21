// Title: Freeze rows and columns and export an Aspose.Cells workbook to a MemoryStream in C#
// Description: Creates a new Workbook, freezes the first two rows and three columns using FreezePanes, saves the workbook to a MemoryStream with SaveToStream, resets the stream position, and writes the stream to a file while handling directory creation and errors.
// Keywords: Aspose.Cells FreezePanes C# | freeze rows columns Aspose.Cells | save workbook to MemoryStream | Aspose.Cells SaveToStream example | export Excel to stream .NET | write MemoryStream to file C# | Aspose.Cells worksheet freeze panes
// Common Searches: Aspose.Cells freeze both rows and columns | How to save Aspose.Cells workbook to MemoryStream | C# example for FreezePanes with Aspose.Cells | Write Aspose.Cells MemoryStream to disk | Export frozen Excel sheet from Aspose.Cells
// Developer Intent: Apply row and column freeze panes to a worksheet and obtain the resulting workbook as a MemoryStream for further processing or file output.
// Use Cases: Generate a report with header rows and columns locked, then stream the file directly to a web client. | Create an Excel attachment in memory, apply FreezePanes, and send it via email without intermediate files. | Build a workbook in a background service, freeze panes for readability, and write the MemoryStream to a dynamically created folder.
// AI Prompts: Show C# code that freezes the first 2 rows and 3 columns in an Aspose.Cells worksheet and returns a MemoryStream. | Demonstrate how to write the MemoryStream from Aspose.Cells.SaveToStream to a file, ensuring the target directory exists. | Explain how to send an Aspose.Cells workbook with frozen panes as an HTTP response using the MemoryStream.

using System;
using System.IO;
using Aspose.Cells;

// Creates a new Workbook, freezes the first two rows and three columns using FreezePanes, saves the workbook to a MemoryStream with SaveToStream, resets the stream position, and writes the stream to a file while handling directory creation and errors.
public class FreezeAndSaveDemo
{
    // Applies row and column freezes and returns the workbook as a memory stream.
    public static MemoryStream Run()
    {
        // Create a new workbook (default format is XLSX).
        Workbook workbook = new Workbook();

        // Get the first worksheet.
        Worksheet worksheet = workbook.Worksheets[0];

        // Freeze the first 2 rows and the first 3 columns.
        // Parameters: row index, column index, number of frozen rows, number of frozen columns.
        worksheet.FreezePanes(2, 3, 2, 3);

        // Save the workbook to a MemoryStream.
        MemoryStream stream = workbook.SaveToStream();

        // Reset the stream position for any subsequent reading.
        stream.Position = 0;

        return stream;
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            using (MemoryStream ms = FreezeAndSaveDemo.Run())
            {
                // Define output file path.
                string outputPath = "FreezeAndSaveDemo.xls";

                // Ensure the directory exists.
                string directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Write the memory stream to a file.
                File.WriteAllBytes(outputPath, ms.ToArray());

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
