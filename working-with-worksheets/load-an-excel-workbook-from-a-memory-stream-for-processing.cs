// Title: Load an Excel workbook from a MemoryStream with Aspose.Cells in C#
// AI Prompts: Write a C# method that accepts a MemoryStream containing an .xlsx file, resets the stream position, validates the argument, and returns an Aspose.Cells Workbook. | Show a console program that reads an Excel file into a byte array, wraps it in a MemoryStream, and creates a Workbook using Aspose.Cells. | Provide error‑handling code that catches any exception thrown while loading a Workbook from a stream and rethrows it as an InvalidOperationException with the original exception attached.
// Common Searches: Aspose.Cells C# load workbook from MemoryStream example | how to open an Excel file from a byte array using Aspose.Cells | resetting MemoryStream position before creating Aspose.Cells Workbook | catching errors when loading Excel workbook from stream with Aspose.Cells
// Tags: load workbook from memory stream Aspose.Cells | read Excel from byte array C# | set MemoryStream.Position to zero for Aspose.Cells | exception handling for Aspose.Cells workbook load | process Excel without saving to disk Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates loading an Excel workbook directly from a MemoryStream using Aspose.Cells, ensuring the stream is positioned at the start and handling errors with a wrapped InvalidOperationException.
public class ExcelProcessor
{
    /// <param name="excelStream">MemoryStream containing the Excel file bytes.</param>
    /// <returns>Workbook object ready for processing.</returns>
    public Workbook LoadWorkbookFromStream(MemoryStream excelStream)
    {
        if (excelStream == null)
            throw new ArgumentNullException(nameof(excelStream));

        // Ensure the stream position is at the beginning before loading
        if (excelStream.CanSeek)
            excelStream.Position = 0;

        try
        {
            // Load the workbook from the memory stream using Aspose.Cells
            Workbook workbook = new Workbook(excelStream);
            return workbook;
        }
        catch (Exception ex)
        {
            // Wrap and rethrow to provide context
            throw new InvalidOperationException("Failed to load workbook from stream.", ex);
        }
    }
}

// Entry point for the console application
public class Program
{
    public static void Main(string[] args)
    {
        const string filePath = "sample.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File not found - {filePath}");
            return;
        }

        try
        {
            // Read the file into a memory stream
            byte[] fileBytes = File.ReadAllBytes(filePath);
            using (MemoryStream ms = new MemoryStream(fileBytes))
            {
                // Load workbook using the helper class
                Workbook workbook = new ExcelProcessor().LoadWorkbookFromStream(ms);
                Console.WriteLine($"Workbook loaded successfully. Worksheets count: {workbook.Worksheets.Count}");
                // Additional processing can be performed here
            }
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
