using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Attempt to modify the application metadata (NameOfApplication)
        SetApplicationMetadata(workbook, "MyCustomApp");

        // Save the workbook (metadata remains unchanged)
        workbook.Save("Output.xlsx");
    }

    // Logs a warning and prevents modification of immutable application metadata
    static void SetApplicationMetadata(Workbook workbook, string newName)
    {
        // Log the warning
        Console.WriteLine($"Warning: Attempt to modify immutable application metadata 'NameOfApplication' to '{newName}' was ignored.");

        // Optionally, display the current value (remains unchanged)
        string current = workbook.BuiltInDocumentProperties.NameOfApplication;
        Console.WriteLine($"Current NameOfApplication remains: '{current}'");
    }
}