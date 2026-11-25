// Title: How to set a worksheet's zoom level to 150% using Aspose.Cells for .NET (C#)
// Common Searches: how to change worksheet zoom to 150 percent with Aspose.Cells C# | Aspose.Cells C# programmatically set Excel sheet view scale | example of adjusting worksheet zoom property in Aspose.Cells .NET | increase Excel worksheet zoom level using Aspose.Cells API
// Tags: set worksheet zoom Aspose.Cells | worksheet zoom property .NET | adjust Excel view scale using Aspose.Cells | Aspose.Cells set zoom level example | save workbook after modifying zoom

using System;
using Aspose.Cells;

// The sample creates a new Workbook, accesses the first Worksheet, sets its Zoom property to 150%, saves the file as ZoomDemo.xlsx, and outputs status messages, all wrapped in a try‑catch block.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (default format is XLSX)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the zoom level to 150%
            worksheet.Zoom = 150;

            // Inform the user
            Console.WriteLine($"Worksheet \"{worksheet.Name}\" zoom set to {worksheet.Zoom}%.");

            // Save the workbook to a file
            string outputFile = "ZoomDemo.xlsx";
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
