using System;
using System.IO;
using Aspose.Cells;

class LoadWorkbookFromStream
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "sample.xlsx";

        try
        {
            // Verify that the file exists before attempting to open it
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"The file '{filePath}' was not found.");

            // Open a read‑only file stream
            using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                // Load the workbook from the stream
                Workbook workbook = new Workbook(stream);

                // Access the first worksheet in the workbook
                Worksheet firstWorksheet = workbook.Worksheets[0];

                // Output the worksheet name
                Console.WriteLine("First worksheet name: " + firstWorksheet.Name);
            }
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine("File error: " + ex.Message);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }
}