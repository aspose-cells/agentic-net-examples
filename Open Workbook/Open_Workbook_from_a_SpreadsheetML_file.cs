using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Define a valid data directory
        string dataDir = Path.Combine(Environment.CurrentDirectory, "Data");
        Directory.CreateDirectory(dataDir);

        // Define input and output file paths
        string inputPath = Path.Combine(dataDir, "sample.xml");
        string outputPath = Path.Combine(dataDir, "converted.xlsx");

        // Ensure the sample XML file exists; create a simple workbook and save as XML if needed
        if (!File.Exists(inputPath))
        {
            Workbook tempWb = new Workbook();
            tempWb.Worksheets[0].Name = "Sheet1";
            tempWb.Save(inputPath, SaveFormat.Xml);
        }

        // Load the workbook from the SpreadsheetML (XML) file
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet and display its name
        Worksheet sheet = workbook.Worksheets[0];
        Console.WriteLine("Worksheet Name: " + sheet.Name);

        // Save the workbook in XLSX format
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}