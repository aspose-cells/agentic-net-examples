// Title: Open an Excel workbook from a byte[] with Aspose.Cells for .NET
// Description: Demonstrates how to wrap a byte array in a MemoryStream, load it into an Aspose.Cells Workbook via the stream constructor, read worksheet information, and save the file as XLSX.
// Keywords: Aspose.Cells load workbook from byte array | C# MemoryStream Excel | Workbook stream constructor | read Excel from byte[] | save workbook after stream load
// Common Searches: Aspose.Cells open workbook from byte array C# | load Excel file using MemoryStream Aspose | convert byte[] to Workbook Aspose.Cells | read first sheet name from byte array Excel
// Developer Intent: Load an Excel file directly from a byte[] without creating an intermediate file, then manipulate or save it.
// Use Cases: Consume Excel data returned by a web service as a byte[] and process it in memory. | Generate an Excel report in memory, store it as a byte array, and later reopen it for further editing. | Retrieve Excel blobs from a database, modify cells, and export to another format using Aspose.Cells.
// AI Prompts: Provide C# code that opens a workbook from a byte[] with Aspose.Cells, iterates all rows, and prints each cell value. | Show how to load an Excel file from a MemoryStream, update a specific cell, and save the workbook as PDF using Aspose.Cells. | Explain strategies for efficiently handling large byte arrays when opening workbooks with Aspose.Cells to minimize memory consumption.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to wrap a byte array in a MemoryStream, load it into an Aspose.Cells Workbook via the stream constructor, read worksheet information, and save the file as XLSX.
public class OpenWorkbookFromByteArrayDemo
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }

    public static void Run()
    {
        // Sample byte array containing an Excel file.
        byte[] excelData = GetSampleExcelData();

        // Wrap the byte array in a MemoryStream.
        using (MemoryStream ms = new MemoryStream(excelData))
        {
            // Ensure the stream position is at the beginning.
            ms.Position = 0;

            // Load the workbook from the stream.
            Workbook workbook = new Workbook(ms);

            // Access the first worksheet and display some information.
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("First worksheet name: " + sheet.Name);
            if (sheet.Cells["A1"].Value != null)
            {
                Console.WriteLine("A1 value: " + sheet.Cells["A1"].StringValue);
            }

            // Save the workbook to a file.
            string outputPath = "LoadedFromByteArray.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved to " + Path.GetFullPath(outputPath));
        }
    }

    // Helper method to create a simple Excel file in memory and return its byte array.
    private static byte[] GetSampleExcelData()
    {
        using (MemoryStream ms = new MemoryStream())
        {
            Workbook temp = new Workbook();
            temp.Worksheets[0].Cells["A1"].PutValue("Hello from byte array");
            temp.Save(ms, SaveFormat.Xlsx);
            return ms.ToArray();
        }
    }
}
