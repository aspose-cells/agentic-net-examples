// Title: C# example: Load an XML file into an Aspose.Cells workbook with LoadOptions and save it as XLSX
// Common Searches: how to load an XML file into an Aspose.Cells workbook using LoadOptions in C# | convert XML data to XLSX with Aspose.Cells in a .NET console application | check for XML file existence before creating an Aspose.Cells workbook in C# | create output folder programmatically when saving an Excel file with Aspose.Cells | save an Aspose.Cells workbook to a specific path as XLSX in C#
// Tags: Aspose.Cells LoadOptions XML import C# | Aspose.Cells save workbook as XLSX .NET | C# verify XML file presence before loading | C# create output directory for Excel export | Aspose.Cells workbook initialization from XML file

using Aspose.Cells;
using System;
using System.IO;

// The sample verifies that a given XML file exists, loads it into an Aspose.Cells Workbook using LoadOptions with LoadFormat.Xml, ensures the target output directory is created, and then saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            const string xmlPath = "data.xml";
            const string outputPath = "output.xlsx";

            // Verify required XML file exists
            if (!File.Exists(xmlPath))
                throw new FileNotFoundException($"XML data file not found: {xmlPath}");

            // Load the XML file into a workbook using appropriate LoadOptions
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xml);
            Workbook workbook = new Workbook(xmlPath, loadOptions);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
