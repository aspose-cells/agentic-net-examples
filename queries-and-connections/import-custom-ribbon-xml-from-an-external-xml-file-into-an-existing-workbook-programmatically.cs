// Title: Programmatically import a custom ribbon XML file into an existing Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load customUI XML from file and embed into existing workbook | How to add a custom ribbon definition to an Excel file using Aspose.Cells .NET | Import external ribbon XML into a workbook programmatically with Aspose.Cells | Saving an Excel workbook with custom ribbon XML using Aspose.Cells in C# | Add custom XML part for ribbon customization in Aspose.Cells workbook
// Tags: add custom ribbon xml part Aspose.Cells | import external customui xml workbook .NET | customxmlparts add ribbon definition C# | embed custom ribbon in Excel using Aspose.Cells | load ribbon xml from file Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example demonstrates loading an existing Excel workbook (or creating a new one), reading a custom ribbon definition from an external XML file, converting the part name and XML content to UTF‑8 byte arrays, adding them as a custom XML part named "customUI" via Workbook.CustomXmlParts.Add, and saving the workbook with the embedded ribbon configuration.
class RibbonImporter
{
    static void Main()
    {
        try
        {
            // Define file paths
            const string workbookPath = "ExistingWorkbook.xlsx";
            const string ribbonPath = "CustomRibbon.xml";
            const string outputPath = "WorkbookWithRibbon.xlsx";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook = File.Exists(workbookPath) ? new Workbook(workbookPath) : new Workbook();

            // Read ribbon XML if the file exists; otherwise use a minimal placeholder
            string ribbonXml = File.Exists(ribbonPath)
                ? File.ReadAllText(ribbonPath)
                : "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\"></customUI>";

            // Convert name and XML to byte arrays as required by the Add method
            byte[] partName = Encoding.UTF8.GetBytes("customUI");
            byte[] partData = Encoding.UTF8.GetBytes(ribbonXml);

            // Add the custom XML part representing the ribbon definition
            workbook.CustomXmlParts.Add(partName, partData);

            // Save the workbook with the imported ribbon
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
