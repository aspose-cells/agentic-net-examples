using System;
using Aspose.Cells;
using System.Text.Json;

class Program
{
    static void Main(string[] args)
    {
        // Path to the workbook; can be passed as a command‑line argument
        string filePath = args.Length > 0 ? args[0] : "sample.xlsm";

        // Load the workbook (Aspose.Cells handles macro‑enabled files)
        Workbook workbook = new Workbook(filePath);

        // Check if the VBA project is signed
        bool isSigned = workbook.VbaProject.IsSigned;

        // If signed, also check whether the signature is valid
        bool isValidSigned = isSigned && workbook.VbaProject.IsValidSigned;

        // Prepare the result object
        var result = new
        {
            IsSigned = isSigned,
            IsValidSigned = isValidSigned
        };

        // Serialize the result to JSON
        string jsonResult = JsonSerializer.Serialize(result);

        // Output the JSON string
        Console.WriteLine(jsonResult);
    }
}