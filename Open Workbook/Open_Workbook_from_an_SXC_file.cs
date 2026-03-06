using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source SXC file
        string sxcFilePath = "sample.sxc";

        // Load the workbook from the SXC file using the string constructor
        Workbook workbook = new Workbook(sxcFilePath);

        // Access the first worksheet
        Worksheet firstSheet = workbook.Worksheets[0];

        // Output basic information about the loaded workbook
        Console.WriteLine("Workbook loaded from SXC file.");
        Console.WriteLine("Number of worksheets: " + workbook.Worksheets.Count);
        Console.WriteLine("First worksheet name: " + firstSheet.Name);

        // Optionally, save the workbook in another format (e.g., XLSX)
        workbook.Save("converted.xlsx", SaveFormat.Xlsx);
    }
}