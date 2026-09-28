// Title: Import XML file into a new Excel workbook and automatically map elements to columns using Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells Workbook.ImportXml to load an XML file and map each element to a separate column starting at cell A1. | Write a C# program that validates the XML file path, creates a fresh workbook, imports the XML content, and writes the result to an XLSX file. | Add logic that creates the target folder if it does not exist before calling Workbook.Save with SaveFormat.Xlsx.
// Common Searches: how to convert an XML document to an Excel sheet with Aspose.Cells in C# | c# automatically map XML nodes to Excel columns using Aspose.Cells ImportXml | save imported Excel workbook to a specific directory with Aspose.Cells .NET
// Tags: aspose.cells importxml to worksheet | c# convert xml elements into spreadsheet columns | create workbook and load xml data | save workbook as xlsx with folder creation | validate xml file existence c#

using System;
using System.IO;
using Aspose.Cells;

// // Checks for the XML file, creates a new Workbook, imports the XML data into Sheet1 starting at A1 (automatically mapping elements to columns), ensures the output folder exists, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the XML file to be imported
            string xmlPath = "data.xml";

            // Verify that the XML file exists to avoid FileNotFoundException
            if (!File.Exists(xmlPath))
                throw new FileNotFoundException($"The XML file '{xmlPath}' was not found.");

            // Create a new workbook (contains a default worksheet named "Sheet1")
            Workbook workbook = new Workbook();

            // Import the XML data into the first worksheet (named "Sheet1") starting at cell A1 (row 0, column 0)
            workbook.ImportXml(xmlPath, "Sheet1", 0, 0);

            // Define output path and ensure its directory exists
            string outputPath = "ImportedData.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook to an Excel file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"XML data successfully imported and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
