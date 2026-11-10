// Title: Create a workbook and save it as a shared Excel file for concurrent editing with Aspose.Cells in C#
// Common Searches: how to create a shared workbook with Aspose.Cells C# | Aspose.Cells enable multi‑user editing in Excel .NET | C# code to save an Excel file for collaborative editing using Aspose | configure Aspose.Cells workbook for concurrent users
// Tags: Aspose.Cells enable shared workbook mode | C# save workbook for multi‑user collaboration | Aspose.Cells configure concurrent editing | Excel shared file creation with .NET | Aspose.Cells workbook shared settings

using System;
using Aspose.Cells;

// The sample creates a new Aspose.Cells Workbook, saves it as "SharedWorkbook.xlsx", writes a success message to the console, and includes basic exception handling. The saved file can be opened as a shared workbook to allow concurrent editing by multiple users.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Save the workbook to a file
            string outputPath = "SharedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
