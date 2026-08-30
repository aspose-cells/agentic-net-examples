// Title: Save an Aspose.Cells workbook to a MemoryStream and then write it as an XLSX file on disk with C#
// AI Prompts: Generate C# code that builds an Aspose.Cells workbook, persists it into a MemoryStream as XLSX, and then saves the stream to a file path. | Demonstrate the correct sequence for positioning a MemoryStream prior to transferring its data to a FileStream when exporting with Aspose.Cells. | Provide an example that changes the output format to PDF while still using a MemoryStream and saving the result to disk. | Explain how to catch and handle exceptions during the process of writing a workbook to a MemoryStream and then persisting it to a file.
// Common Searches: aspocells c# export workbook using memory stream and file | how to export an Excel workbook as xlsx using a MemoryStream in Aspose.Cells | reset memory stream position before copying to file stream c# | write Aspose.Cells output from memory stream to disk | persist Aspose.Cells workbook via memory stream and write to server file c#
// Tags: save workbook to memory stream Aspose.Cells | write memory stream to file C# | export workbook as xlsx via stream Aspose.Cells | memory stream position handling C# | copy stream to file stream Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Aspose.Cells workbook, adds sample data, saves it to a MemoryStream in XLSX format, resets the stream position, and copies the stream to a file named OutputFromStream.xlsx on disk.
public class SaveWorkbookToMemoryAndDisk
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook
        using (Workbook workbook = new Workbook())
        {
            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Save the workbook to a memory stream in XLSX format
            using (MemoryStream memoryStream = new MemoryStream())
            {
                workbook.Save(memoryStream, SaveFormat.Xlsx);

                // Reset the stream position before reading
                memoryStream.Position = 0;

                // Write the stream content to a file on disk
                string outputPath = "OutputFromStream.xlsx";
                using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    memoryStream.CopyTo(fileStream);
                }
            }
        }
    }
}
