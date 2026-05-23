using System;
using Aspose.Cells;

class CreateWorkbookScopedNamedRange
{
    static void Main()
    {
        try
        {
            // Initialize a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Create the range A1:D10 on the first worksheet
            // Use fully qualified type to avoid conflict with System.Range
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1", "D10");

            // Add a new name to the workbook's name collection (global scope)
            int nameIndex = workbook.Worksheets.Names.Add("MyRange");

            // Retrieve the created Name object
            Name namedRange = workbook.Worksheets.Names[nameIndex];

            // Define the reference for the name (workbook‑scoped)
            // SheetIndex = 0 indicates a global name (workbook scope)
            namedRange.SheetIndex = 0;
            namedRange.RefersTo = $"={sheet.Name}!$A$1:$D$10";

            // Optionally, associate the name with the Range object (not required for scope)
            range.Name = "MyRange";

            // Save the workbook to a file
            string outputPath = "WorkbookScopedNamedRange.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}