using System;
using Aspose.Cells;
using System.Text.Json;

class Program
{
    static void Main(string[] args)
    {
        // Input Excel file (macro-enabled). Use first argument if provided, otherwise default.
        string inputPath = args.Length > 0 ? args[0] : "sample.xlsm";

        // Load the workbook (Aspose.Cells handles macro-enabled files automatically).
        Workbook workbook = new Workbook(inputPath);

        // Retrieve the VBA project protection flag.
        bool isProtected = workbook.VbaProject.IsProtected;

        // Prepare an anonymous object for JSON serialization.
        var status = new { IsVbaProjectProtected = isProtected };

        // Convert the object to a JSON string.
        string jsonResult = JsonSerializer.Serialize(status);

        // Output the JSON result.
        Console.WriteLine(jsonResult);
    }
}