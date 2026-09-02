// Title: Load an Excel workbook, set author, title, subject, and keywords via reflection, update a cell, and save as XLSX with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an existing .xlsx file with Aspose.Cells, uses reflection to assign Author, Title, Subject, and Keywords properties, changes the value of cell A1, and saves the workbook to a new .xlsx file. | Show how to safely handle errors when loading a workbook, setting document properties through reflection, and saving the file using Aspose.Cells SaveFormat.Xlsx. | Demonstrate accessing the WorkbookProperties object via reflection to maintain compatibility across different Aspose.Cells versions in a .NET application.
// Common Searches: aspnet set workbook author and keywords using aspose.cells reflection | how to modify cell A1 and preserve document properties when saving Excel with Aspose.Cells | load existing xlsx, change metadata, and save as new file using Aspose.Cells for .NET | reflection based access to WorkbookProperties in Aspose.Cells C# example | error handling for loading and saving Excel files with Aspose.Cells
// Tags: Aspose.Cells set workbook metadata via reflection | C# modify Excel cell and document properties | save workbook as XLSX using Aspose.Cells SaveFormat | load existing workbook with fallback creation Aspose.Cells | compatible WorkbookProperties access across Aspose.Cells versions

using Aspose.Cells;
using System;
using System.IO;
using System.Reflection;

// The example loads an existing XLSX workbook (or creates a new one if missing), uses reflection to set Author, Title, Subject, and Keywords in the workbook's properties, updates cell A1 with a custom value, and saves the modified workbook as a new XLSX file while handling potential errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook if it exists; otherwise create a new one
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load '{inputPath}': {ex.Message}");
                    workbook = new Workbook();
                }
            }
            else
            {
                workbook = new Workbook();
            }

            // Set document properties via reflection (compatible with all Aspose.Cells versions)
            try
            {
                PropertyInfo wpInfo = workbook.GetType().GetProperty("WorkbookProperties");
                if (wpInfo != null)
                {
                    object props = wpInfo.GetValue(workbook);
                    Type propsType = props.GetType();

                    propsType.GetProperty("Author")?.SetValue(props, "John Doe");
                    propsType.GetProperty("Title")?.SetValue(props, "Modified Workbook");
                    propsType.GetProperty("Subject")?.SetValue(props, "Aspose.Cells Demo");
                    propsType.GetProperty("Keywords")?.SetValue(props, "Aspose, Cells, C#");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unable to set document properties: {ex.Message}");
            }

            // Example modification: change the value of cell A1 in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Hello, Aspose!");

            // Save the workbook to XLSX format
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
