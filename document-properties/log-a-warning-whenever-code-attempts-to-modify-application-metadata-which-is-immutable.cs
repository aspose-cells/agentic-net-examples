// Title: Log a warning in C# when attempting to modify immutable built‑in document properties with Aspose.Cells
// Common Searches: C# Aspose.Cells log warning for attempted change of built‑in document property | how to prevent modification of Excel built‑in properties using Aspose.Cells .NET | read‑only document properties handling in Aspose.Cells workbook | Aspose.Cells .NET warning when setting Author property on workbook | detect and log immutable metadata changes in Aspose.Cells C#
// Tags: log warning on immutable document property Aspose.Cells | prevent modification of built‑in workbook metadata C# | Aspose.Cells read‑only built‑in properties handling | C# workbook property immutability Aspose.Cells | Excel metadata immutability Aspose.Cells .NET

using System;
using Aspose.Cells;

// The example creates a Workbook, calls a helper method that logs a warning whenever code tries to change built‑in document properties such as Author or Title, leaves the properties unchanged because they are immutable, and then saves the workbook to an XLSX file.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Attempt to modify application metadata.
        // In this scenario we treat built‑in document properties as immutable.
        // Instead of changing them directly, we log a warning.
        SetImmutableMetadata(workbook, "Author", "John Doe");
        SetImmutableMetadata(workbook, "Title", "Sample Report");

        // Save the workbook (lifecycle rule)
        workbook.Save("output.xlsx");
    }

    // Helper method that logs a warning whenever code tries to modify metadata.
    // The actual modification is intentionally omitted because the metadata is immutable.
    static void SetImmutableMetadata(Workbook wb, string propertyName, string newValue)
    {
        // Log a warning indicating that an attempt was made to change immutable metadata.
        Console.WriteLine($"Warning: Attempt to modify immutable application metadata '{propertyName}'. Change ignored.");
        
        // If you still need to read the current value, you can do so safely:
        // Example for built‑in properties:
        // string currentValue = wb.BuiltInDocumentProperties[propertyName]?.ToString();
        // Console.WriteLine($"Current value of '{propertyName}': {currentValue}");
    }
}
