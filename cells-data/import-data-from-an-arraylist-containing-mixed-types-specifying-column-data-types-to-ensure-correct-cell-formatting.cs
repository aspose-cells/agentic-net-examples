// Title: Import a Mixed‑Type ArrayList into Excel with Aspose.Cells (C#)
// Description: Creates a Workbook, converts an ArrayList that holds a string, int, double, DateTime and bool into an object array, and uses Cells.ImportObjectArray to write the values horizontally starting at A1 while preserving each data type. The workbook is then saved as an XLSX file.
// Keywords: Aspose.Cells | ImportObjectArray | ArrayList to Excel | mixed data types | C# Excel export | preserve data types | DateTime cell formatting | boolean Excel value | numeric formatting | Excel worksheet import
// Common Searches: Aspose.Cells import ArrayList mixed types C# | ImportObjectArray preserve DateTime and bool | How to export heterogeneous collection to Excel .NET | C# write string, int, double, DateTime to Excel in one row
// Developer Intent: Write a heterogeneous ArrayList to an Excel worksheet while keeping each element’s native type.
// Use Cases: Export a product catalog row (name, quantity, price, expiry date, in‑stock flag) in a single call. | Generate daily status reports that include dates and boolean flags directly from a collection. | Populate a template row with mixed‑type test data for QA or demo purposes.
// AI Prompts: Show C# code that converts a List<object> containing strings, numbers, dates and booleans into an object array and imports it into a specific Excel range using Aspose.Cells ImportObjectArray, then applies appropriate number and date formats. | Explain how to use Aspose.Cells ImportObjectArray to write mixed‑type data to a worksheet while preserving each type, and how to save the workbook as XLSX.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a Workbook, converts an ArrayList that holds a string, int, double, DateTime and bool into an object array, and uses Cells.ImportObjectArray to write the values horizontally starting at A1 while preserving each data type. The workbook is then saved as an XLSX file.
    public class ImportArrayListWithMixedTypes
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet's cells collection
                Workbook workbook = new Workbook();
                Cells cells = workbook.Worksheets[0].Cells;

                // Prepare an ArrayList containing mixed data types
                ArrayList mixedData = new ArrayList
                {
                    "Product",                     // string
                    150,                           // integer
                    99.99,                         // double
                    new DateTime(2023, 12, 31),    // DateTime
                    true                           // boolean
                };

                // Convert the ArrayList to an object array – ImportObjectArray preserves each value's type
                object[] dataArray = (object[])mixedData.ToArray(typeof(object));

                // Import the data horizontally starting at cell A1 (row 0, column 0)
                cells.ImportObjectArray(dataArray, 0, 0, false);

                // Define output file path
                string outputPath = "ImportArrayListMixedTypes.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ImportArrayListWithMixedTypes.Run();
        }
    }
}
