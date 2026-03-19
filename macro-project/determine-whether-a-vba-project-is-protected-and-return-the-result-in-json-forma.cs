using System;
using Aspose.Cells;
using Aspose.Cells.Vba;
using System.Text.Json;

public class VbaProjectProtectionChecker
{
    public static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Please provide the path to the Excel file as an argument.");
            return;
        }

        Run(args[0]);
    }

    // Checks if the VBA project in the given workbook is protected and outputs JSON.
    public static void Run(string filePath)
    {
        // Load the workbook from the specified file.
        Workbook workbook = new Workbook(filePath);

        // Access the VBA project associated with the workbook.
        VbaProject vbaProject = workbook.VbaProject;

        // If there is no VBA project, consider it not protected.
        bool isProtected = vbaProject != null && vbaProject.IsProtected;

        // Create an anonymous object to hold the result.
        var result = new { IsProtected = isProtected };

        // Serialize the result to JSON format.
        string json = JsonSerializer.Serialize(result);

        // Write the JSON string to the console.
        Console.WriteLine(json);
    }
}