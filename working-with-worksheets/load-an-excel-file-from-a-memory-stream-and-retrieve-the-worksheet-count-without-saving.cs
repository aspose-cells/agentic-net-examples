// Title: Count worksheets in an Excel file loaded from a byte array using Aspose.Cells for .NET
// AI Prompts: Write a C# method that receives a byte[] of an Excel workbook, creates a MemoryStream, loads it with Aspose.Cells, and returns the total number of worksheets. | Demonstrate how to read an Excel file into a byte array, pass it to a helper class, and print the worksheet count without writing the file to disk using Aspose.Cells.
// Common Searches: aspnet count worksheets in excel file from memory stream | aspocells get number of sheets from byte array | c# load excel workbook from byte[] without saving to disk | how to retrieve worksheet count using Aspose.Cells and MemoryStream | read excel file into byte array and count sheets in .NET
// Tags: Aspose.Cells load workbook from byte array | retrieve worksheet count in .NET | C# MemoryStream Excel workbook | count sheets without saving file | Aspose.Cells worksheets collection count

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook from a byte array via MemoryStream using Aspose.Cells and returns the total number of worksheets in the workbook.
public class ExcelHelper
{
    /// <param name="excelData">Byte array containing the Excel file.</param>
    /// <returns>Count of worksheets in the workbook.</returns>
    public int GetWorksheetCount(byte[] excelData)
    {
        // Create a memory stream from the provided byte array
        using (MemoryStream stream = new MemoryStream(excelData))
        {
            // Load the workbook from the memory stream
            Workbook workbook = new Workbook(stream);

            // Retrieve and return the worksheet count
            return workbook.Worksheets.Count;
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: AsposeCellsRunner <excel-file-path>");
                return;
            }

            string filePath = args[0];

            // Prevent FileNotFoundException by checking existence
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File not found - {filePath}");
                return;
            }

            // Read the Excel file into a byte array
            byte[] excelData = File.ReadAllBytes(filePath);

            // Use ExcelHelper to get worksheet count
            ExcelHelper helper = new ExcelHelper();
            int sheetCount = helper.GetWorksheetCount(excelData);

            Console.WriteLine($"Worksheet count: {sheetCount}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
