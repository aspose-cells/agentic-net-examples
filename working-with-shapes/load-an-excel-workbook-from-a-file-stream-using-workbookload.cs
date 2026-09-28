// Title: Load an Excel workbook from a FileStream using Aspose.Cells Workbook constructor in C#
// AI Prompts: Write a C# method that takes a file path, checks that the file exists, opens a FileStream for read access, and creates an Aspose.Cells Workbook from the stream while wrapping any errors in a clear exception. | Generate C# sample code that demonstrates loading a .xlsx file into an Aspose.Cells Workbook using a FileStream inside a using block.
// Common Searches: Aspose.Cells C# load workbook from FileStream example | How to open an Excel file with Aspose.Cells using a stream in .NET | C# read .xlsx file with Aspose.Cells and handle missing file errors
// Tags: Aspose.Cells workbook load from stream C# | C# FileStream Excel loading Aspose.Cells | exception handling Aspose.Cells workbook creation | validate file existence before loading Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The ExcelLoader class provides a LoadWorkbookFromStream method that verifies the file's existence, opens it with a FileStream in read mode, constructs an Aspose.Cells Workbook from that stream, and returns the workbook while encapsulating any failures in an InvalidOperationException.
    public class ExcelLoader
    {
        // Loads an Excel workbook from a file stream using the Workbook constructor
        public Workbook LoadWorkbookFromStream(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"The file '{filePath}' was not found.", filePath);

            try
            {
                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    // Load workbook directly from the stream
                    Workbook workbook = new Workbook(stream);
                    return workbook;
                }
            }
            catch (Exception ex)
            {
                // Wrap and rethrow for caller handling
                throw new InvalidOperationException("Failed to load workbook from stream.", ex);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage: provide path to an existing Excel file
            string path = "sample.xlsx";

            try
            {
                ExcelLoader loader = new ExcelLoader();
                Workbook wb = loader.LoadWorkbookFromStream(path);
                Console.WriteLine($"Workbook loaded successfully. Worksheets count: {wb.Worksheets.Count}");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine(fnfEx.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
