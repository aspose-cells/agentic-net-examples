// Title: How to save a modified Excel workbook as a new file while keeping the original formatting using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save workbook to new file without losing cell styles | preserve original Excel formatting when saving a modified workbook with Aspose.Cells | copy Excel workbook and export as separate file while retaining formatting in .NET | C# Aspose.Cells keep existing styles after modifying and saving workbook | use SaveFormat.Xlsx to retain formatting when saving changes to Excel
// Tags: Aspose.Cells save workbook preserving formatting | C# export modified Excel as new file with styles | SaveFormat.Xlsx retain original cell styles | duplicate workbook without losing formatting Aspose | preserve Excel formatting on save .NET

using Aspose.Cells;

// The example loads an existing Excel file (Original.xlsx) with Aspose.Cells, allows any modifications, and then saves the workbook to a new file (Modified.xlsx) using SaveFormat.Xlsx, ensuring that all original formatting and styles are retained.
class Program
{
    static void Main()
    {
        // Load the original workbook (original formatting is retained)
        Workbook workbook = new Workbook("Original.xlsx");

        // TODO: apply any modifications to the workbook here

        // Save the modified workbook to a new file while preserving formatting
        workbook.Save("Modified.xlsx", SaveFormat.Xlsx);
    }
}
