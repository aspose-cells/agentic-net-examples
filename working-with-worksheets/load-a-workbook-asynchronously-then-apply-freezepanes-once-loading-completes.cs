using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

class Program
{
    static async Task Main(string[] args)
    {
        // Paths for input and output files
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Load the workbook asynchronously
        Workbook workbook = await LoadWorkbookAsync(inputPath);

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Freeze panes at cell C3 with 3 frozen rows and 3 frozen columns
        sheet.FreezePanes("C3", 3, 3);

        // Save the modified workbook
        workbook.Save(outputPath);
    }

    // Asynchronously reads a file into a MemoryStream and creates a Workbook from it
    static async Task<Workbook> LoadWorkbookAsync(string filePath)
    {
        // Read all bytes from the file asynchronously
        byte[] fileBytes = await File.ReadAllBytesAsync(filePath);

        // Create a MemoryStream from the byte array
        using (MemoryStream stream = new MemoryStream(fileBytes))
        {
            // Use the Workbook(Stream) constructor (provided rule) to load the workbook
            Workbook wb = new Workbook(stream);
            return wb;
        }
    }
}