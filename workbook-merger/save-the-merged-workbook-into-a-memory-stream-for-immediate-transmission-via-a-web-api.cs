// Title: Merge Excel workbooks and return the result as a MemoryStream using Aspose.Cells for .NET
// Description: C# example that validates two XLSX files, loads them with Aspose.Cells, uses the Combine method to merge the source workbook into the destination workbook, saves the merged workbook to a MemoryStream in XLSX format, resets the stream position, and returns the stream for instant use in a web API or micro‑service.
// Keywords: Aspose.Cells | C# | .NET | merge workbooks | Combine method | MemoryStream | in‑memory Excel | XLSX stream | web API response | file download | streaming Excel
// Common Searches: Aspose.Cells merge two workbooks to memory stream | C# combine Excel files and return stream | save merged workbook to MemoryStream for API | Aspose.Cells Combine example .NET | return Excel file as byte array using Aspose
// Developer Intent: Create a merged Excel workbook from two files and obtain it as a MemoryStream for immediate transmission via an API.
// Use Cases: Combine user‑uploaded Excel reports on a server and send the merged file back as an HTTP response without writing to disk. | Aggregate monthly financial workbooks in a cloud function, keep the result in memory, and pass it to downstream services. | Implement a microservice that merges multiple spreadsheets and returns the result as a byte array for further processing.
// AI Prompts: Generate C# code that merges several workbooks with Aspose.Cells and returns the combined file as a MemoryStream for an ASP.NET Core file download. | Show how to modify the method to output CSV instead of XLSX while still returning a MemoryStream. | Explain strategies for handling very large Excel files during merging and streaming with Aspose.Cells to minimize memory consumption.

using System;
using System.IO;
using Aspose.Cells;

// C# example that validates two XLSX files, loads them with Aspose.Cells, uses the Combine method to merge the source workbook into the destination workbook, saves the merged workbook to a MemoryStream in XLSX format, resets the stream position, and returns the stream for instant use in a web API or micro‑service.
public class WorkbookMergeService
{
    /// <param name="firstFilePath">Full path to the first workbook.</param>
    /// <param name="secondFilePath">Full path to the second workbook.</param>
    /// <returns>A MemoryStream containing the merged workbook in XLSX format.</returns>
    public MemoryStream MergeWorkbooksAndGetStream(string firstFilePath, string secondFilePath)
    {
        try
        {
            // Verify that both input files exist
            if (!File.Exists(firstFilePath))
                throw new FileNotFoundException($"The file '{firstFilePath}' was not found.", firstFilePath);
            if (!File.Exists(secondFilePath))
                throw new FileNotFoundException($"The file '{secondFilePath}' was not found.", secondFilePath);

            // Load the destination workbook (the one that will receive the other workbook's content)
            Workbook destWorkbook = new Workbook(firstFilePath);

            // Load the source workbook (the one to be combined into the destination)
            Workbook sourceWorkbook = new Workbook(secondFilePath);

            // Combine the source workbook into the destination workbook
            destWorkbook.Combine(sourceWorkbook);

            // Create a memory stream to hold the merged workbook
            MemoryStream mergedStream = new MemoryStream();

            // Save the combined workbook into the memory stream using XLSX format
            destWorkbook.Save(mergedStream, SaveFormat.Xlsx);

            // Reset the stream position so that consumers can read from the beginning
            mergedStream.Position = 0;

            // Return the prepared stream
            return mergedStream;
        }
        catch (Exception ex)
        {
            // Log or rethrow as needed; here we wrap in an ApplicationException for clarity
            throw new ApplicationException("An error occurred while merging workbooks.", ex);
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Example usage:
        // Provide two existing Excel file paths as command‑line arguments or modify the paths below.
        string firstPath = args.Length > 0 ? args[0] : "FirstWorkbook.xlsx";
        string secondPath = args.Length > 1 ? args[1] : "SecondWorkbook.xlsx";

        try
        {
            var service = new WorkbookMergeService();
            using (MemoryStream mergedStream = service.MergeWorkbooksAndGetStream(firstPath, secondPath))
            {
                // Save the merged stream to a file for verification
                string outputPath = "MergedWorkbook.xlsx";
                using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    mergedStream.CopyTo(fileStream);
                }
                Console.WriteLine($"Merged workbook saved to '{outputPath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
