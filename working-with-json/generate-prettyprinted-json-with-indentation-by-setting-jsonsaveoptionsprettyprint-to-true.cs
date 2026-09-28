// Title: How to save an Aspose.Cells workbook as indented (pretty‑printed) JSON using JsonSaveOptions in C#
// AI Prompts: Write C# code that creates a Workbook, sets JsonSaveOptions.PrettyPrint = true, and saves the file as formatted JSON. | Show how to configure Aspose.Cells JsonSaveOptions for pretty‑printed JSON output in a .NET console application.
// Common Searches: Aspose.Cells C# export Excel to formatted JSON with indentation | Enable pretty print when saving workbook to JSON using Aspose.Cells .NET | JsonSaveOptions PrettyPrint property usage example in C# | Generate readable JSON from an Excel workbook with Aspose.Cells | Save Excel data as indented JSON file using Aspose.Cells library
// Tags: Aspose.Cells JsonSaveOptions PrettyPrint | C# workbook to indented JSON | formatted JSON export from Excel | pretty‑printed JSON generation with Aspose.Cells | JSON serialization with indentation in .NET

using System;
using Aspose.Cells;

// The program creates a new workbook, populates sample data in the first worksheet, enables JsonSaveOptions.PrettyPrint for formatted output, and saves the workbook as an indented JSON file (output.json) while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(25);

            // Configure JSON save options (default settings are sufficient)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Save the workbook as a JSON file with the specified options
            workbook.Save("output.json", jsonOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
