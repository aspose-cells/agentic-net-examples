// Title: Import an XML map from a file into an Aspose.Cells workbook using C# Workbook.ImportXml
// Common Searches: C# Aspose.Cells import XML file as a map into a new workbook | Using Workbook.ImportXml to add XML schema to an Excel file with Aspose.Cells | Aspose.Cells example for loading XML data and creating an XML map in C# | How to check XML file existence before calling Workbook.ImportXml in C# | Save workbook to XLSX after importing an XML map with Aspose.Cells
// Tags: Workbook.ImportXml XML map import | Aspose.Cells import XML schema C# | Create XML map from file Aspose.Cells | Save workbook to XLSX after XML import | Validate XML file existence C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, checks that the source XML file exists, imports the XML file (including its schema) as a map named "MyXmlMap" using Workbook.ImportXml with default flags, saves the workbook as Result.xlsx, and writes success or error messages to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Path to the source XML file
            string xmlPath = "data.xml";

            // Verify that the XML file exists before attempting import
            if (!File.Exists(xmlPath))
            {
                Console.WriteLine($"Error: XML file not found at path '{xmlPath}'.");
                return;
            }

            try
            {
                // Import XML data (and schema) into the workbook.
                // Note: In some Aspose.Cells versions the flags are integers (0 = false, 1 = true).
                workbook.ImportXml(xmlPath, "MyXmlMap", 0, 0);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during XML import: {ex.Message}");
                return;
            }

            // Save the workbook to an Excel file
            string resultPath = "Result.xlsx";
            workbook.Save(resultPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
