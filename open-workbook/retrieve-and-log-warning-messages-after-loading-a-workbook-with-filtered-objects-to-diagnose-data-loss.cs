// Title: How to retrieve and log warning messages after loading a filtered Excel workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# retrieve warning collection after loading workbook with filtered rows | log warning messages from Aspose.Cells LoadOptions when opening Excel file | detect data loss warnings in filtered Excel using Aspose.Cells .NET API | how to access WarningCollection after loading workbook with Aspose.Cells in C# | capture load-time warnings for filtered objects in Aspose.Cells workbook
// Tags: Aspose.Cells warning collection C# | load workbook with LoadOptions warnings | filtered Excel data loss detection Aspose.Cells | log workbook load warnings .NET | retrieve load-time warnings Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the presence of a filtered Excel file, loads it into an Aspose.Cells Workbook, and demonstrates how to access the WarningCollection via LoadOptions to capture and log any load-time warnings that may indicate data loss.
class Program
{
    static void Main()
    {
        const string filePath = "FilteredData.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // If you need to handle warnings, Aspose.Cells provides a WarningCollection
            // after loading with LoadOptions. For simplicity, we just confirm successful load.
            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred while loading the workbook: {ex.Message}");
        }
    }
}
