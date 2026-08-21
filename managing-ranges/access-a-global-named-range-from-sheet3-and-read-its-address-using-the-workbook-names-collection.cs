// Title: Aspose.Cells C# – Retrieve Address and Worksheet of a Global Named Range on Sheet3
// Description: Shows how to create a workbook, add a sheet named Sheet3, define a global named range (MyGlobalRange) that points to Sheet3!A1:B3, and then use the workbook's Names collection to obtain the Range object, read its address and parent worksheet name, and save the file.
// Keywords: Aspose.Cells | C# | .NET | global named range | Names collection | GetRange | range address | worksheet name | Sheet3 | Excel automation
// Common Searches: Aspose.Cells get address of global named range | C# retrieve worksheet of named range using Aspose.Cells | How to read a named range address from workbook Names collection | Aspose.Cells example for global named range on specific sheet | GetRange address Sheet3 Aspose.Cells .NET
// Developer Intent: Obtain the address and owning worksheet of a global named range defined on Sheet3 via the Aspose.Cells Names collection.
// Use Cases: Debugging: display the address of a global named range in a console app. | Validation: confirm that a named range points to the expected worksheet. | Pre‑processing: read range metadata before performing calculations or exports.
// AI Prompts: Generate C# code with Aspose.Cells that creates a global named range on Sheet3 and prints its address and worksheet name. | Explain how to fetch a named range from the workbook's Names collection and retrieve its Range object in Aspose.Cells. | Show a step‑by‑step example of reading the address of a global named range and identifying its parent sheet using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add a sheet named Sheet3, define a global named range (MyGlobalRange) that points to Sheet3!A1:B3, and then use the workbook's Names collection to obtain the Range object, read its address and parent worksheet name, and save the file.
    public class AccessGlobalNamedRangeFromSheet3
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a worksheet named "Sheet3"
                Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

                // Populate some data in Sheet3 (optional, just for demonstration)
                sheet3.Cells["A1"].PutValue("Item");
                sheet3.Cells["B1"].PutValue("Quantity");
                sheet3.Cells["A2"].PutValue("Apple");
                sheet3.Cells["B2"].PutValue(10);
                sheet3.Cells["A3"].PutValue("Orange");
                sheet3.Cells["B3"].PutValue(15);

                // Create a global named range that refers to a range on Sheet3
                NameCollection names = workbook.Worksheets.Names;
                int nameIndex = names.Add("MyGlobalRange");
                Name globalName = names[nameIndex];
                // Set the reference to Sheet3!A1:B3
                globalName.RefersTo = "=Sheet3!$A$1:$B$3";

                // Access the global named range using its text
                Name retrievedName = names["MyGlobalRange"]; // using the string indexer
                // Get the Range object that the name refers to
                AsposeRange range = retrievedName.GetRange();

                // Read and display the address of the range
                Console.WriteLine($"Global named range address: {range.Address}");
                Console.WriteLine($"Worksheet of the range: {range.Worksheet.Name}");

                // Save the workbook (optional)
                string outputPath = "AccessGlobalNamedRangeFromSheet3.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AccessGlobalNamedRangeFromSheet3.Run();
        }
    }
}
