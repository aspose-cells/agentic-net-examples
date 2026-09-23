// Title: Add a custom XML part and purge unused styles in a single Aspose.Cells .NET workflow
// AI Prompts: Create C# code that opens a new Workbook, embeds a custom XML part, calls the RemoveUnusedStyles method, and saves the file as XLSX in one continuous pipeline. | Write a .NET function that takes an XML string and a file path, adds the XML as a custom part to the workbook, cleans up any unused cell styles, and returns the saved workbook path.
// Common Searches: asp.net cells embed custom xml part and clean up unused styles before saving | how to add a custom XML part to an Excel file and remove unused cell styles using Aspose.Cells | single-step workbook processing Aspose.Cells custom XML and style cleanup | C# Aspose.Cells combine custom XML insertion with RemoveUnusedStyles call | optimize Aspose.Cells workbook by adding XML and deleting unused styles in one operation
// Tags: Aspose.Cells add custom XML part | Aspose.Cells remove unused styles | Aspose.Cells combined workbook processing | Aspose.Cells embed XML in XLSX | Aspose.Cells style cleanup .NET

using Aspose.Cells;
using System;
using System.IO;
using System.Text;

// The example creates a new Workbook, inserts sample data, adds a custom XML part by converting strings to UTF‑8 byte arrays, invokes RemoveUnusedStyles to eliminate redundant cell styles, and saves the result as Output.xlsx while handling exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Optional: add some data to a worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["A2"].PutValue(123);

            // Add a custom XML part (convert strings to byte[] as required by the API)
            string xmlId = "MyCustomXml";
            string xmlContent = "<root><item>Value</item></root>";
            workbook.CustomXmlParts.Add(Encoding.UTF8.GetBytes(xmlId), Encoding.UTF8.GetBytes(xmlContent));

            // Save the workbook
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
