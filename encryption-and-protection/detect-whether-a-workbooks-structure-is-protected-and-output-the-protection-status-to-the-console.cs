// Title: Check if an Excel workbook's structure is protected using Aspose.Cells for .NET and output the result
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, reads the Workbook.Protection.IsStructureProtected flag, and writes the boolean value to the console. | Demonstrate how to use reflection in C# to obtain the IsStructureProtected property from the Workbook.Protection object when the property is not directly accessible. | Create a robust console program that reports workbook structure protection, includes file‑existence validation, and provides a fallback message for older Aspose.Cells versions lacking the Protection API.
// Common Searches: aspocells c# determine if workbook structure is locked | how to read IsStructureProtected property from an Excel file using Aspose.Cells | C# reflection to access Workbook protection information in Aspose.Cells | fallback method for workbook protection status in older Aspose.Cells releases | console output workbook structure protection status Aspose.Cells .NET
// Tags: check workbook structure protection Aspose.Cells | read Workbook.Protection.IsStructureProtected via reflection | fallback for Aspose.Cells versions without Protection property | C# console report Excel protection status | validate input file existence Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, uses reflection to retrieve the Workbook.Protection.IsStructureProtected flag when available, prints the protection status to the console, validates the input file, and supplies a fallback message for older library versions that lack the Protection property.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Attempt to retrieve workbook protection information via reflection
            PropertyInfo protectionProp = typeof(Workbook).GetProperty("Protection");
            if (protectionProp != null)
            {
                object protectionObj = protectionProp.GetValue(workbook);
                PropertyInfo isStructureProp = protectionObj?.GetType().GetProperty("IsStructureProtected");
                bool isStructureProtected = false;

                if (isStructureProp != null && isStructureProp.PropertyType == typeof(bool))
                {
                    isStructureProtected = (bool)isStructureProp.GetValue(protectionObj);
                }

                Console.WriteLine("Workbook structure protected: " + isStructureProtected);
            }
            else
            {
                // Fallback message for older Aspose.Cells versions lacking the Protection property
                Console.WriteLine("Workbook protection information not available in this Aspose.Cells version.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
