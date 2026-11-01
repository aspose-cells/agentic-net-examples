// Title: Check if an Excel worksheet allows editing objects with Aspose.Cells for .NET and log the result
// Common Searches: how to determine if a worksheet's AllowEditObject flag is true using Aspose.Cells in C# | C# read worksheet protection AllowEditObject property with Aspose.Cells | log worksheet protection settings for editing objects in Aspose.Cells .NET | use reflection to get worksheet protection AllowEditObject in Aspose.Cells | detect editing objects permission in an Excel file using Aspose.Cells API
// Tags: Aspose.Cells worksheet protection AllowEditObject | C# read worksheet protection flag | log worksheet editing objects permission | reflection access Aspose.Cells protection property | detect editing objects flag Excel .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells for .NET, accesses the first worksheet, and uses reflection to read the AllowEditObject flag from the worksheet's Protection object. It then writes a console message indicating whether editing objects is permitted, handling missing files and reflection errors gracefully.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve protection settings via reflection to avoid direct type dependency
            bool allowsEditObjects = false;
            try
            {
                object protectionObj = worksheet.Protection;
                var propInfo = protectionObj.GetType().GetProperty("AllowEditObject");
                if (propInfo != null)
                {
                    allowsEditObjects = (bool)propInfo.GetValue(protectionObj);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to read protection settings. {ex.Message}");
            }

            // Log the compliance result
            if (allowsEditObjects)
            {
                Console.WriteLine($"Worksheet '{worksheet.Name}' allows editing objects.");
            }
            else
            {
                Console.WriteLine($"Worksheet '{worksheet.Name}' does NOT allow editing objects.");
            }

            // No modifications made; saving is optional
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
