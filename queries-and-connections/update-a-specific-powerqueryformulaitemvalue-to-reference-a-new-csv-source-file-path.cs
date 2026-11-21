// Title: Update PowerQueryFormulaItem.Value to point to a new CSV source file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change PowerQueryFormulaItem value to new CSV path | How to modify Power Query source file in an existing Excel workbook with Aspose.Cells | Set Csv.Document source in PowerQueryFormulaItems via reflection Aspose.Cells .NET | Update Power Query formula for a named query using Aspose.Cells API | Replace CSV file reference in Power Query connection of an XLSX using Aspose.Cells C#
// Tags: set PowerQueryFormulaItem formula | Aspose.Cells change Power Query source path | C# modify Power Query CSV reference | Excel workbook Power Query formula edit | reflection access PowerQueryFormulaItems Aspose.Cells

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, obtains the PowerQueryFormulaItems collection via reflection to support multiple API versions, finds the item named "MyQuery", replaces its Value with a new Csv.Document formula that points to "C:\\NewFolder\\newfile.csv", and saves the modified workbook.
class UpdatePowerQueryFormula
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Try to obtain Power Query items via reflection (covers different API versions)
            object pqItemsObj = null;
            var pqProp = worksheet.GetType().GetProperty("PowerQueryFormulaItems");
            if (pqProp != null)
            {
                pqItemsObj = pqProp.GetValue(worksheet);
            }
            else
            {
                var pqMethod = worksheet.GetType().GetMethod("GetPowerQueryFormulaItems");
                if (pqMethod != null)
                {
                    pqItemsObj = pqMethod.Invoke(worksheet, null);
                }
            }

            if (pqItemsObj == null)
            {
                Console.WriteLine("Power Query functionality is not available in this Aspose.Cells version.");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
                return;
            }

            // Iterate through the collection to find the target item
            bool itemFound = false;
            foreach (object item in (IEnumerable)pqItemsObj)
            {
                var nameProp = item.GetType().GetProperty("Name");
                var valueProp = item.GetType().GetProperty("Value");
                if (nameProp == null || valueProp == null) continue;

                string name = nameProp.GetValue(item) as string;
                if (string.Equals(name, "MyQuery", StringComparison.OrdinalIgnoreCase))
                {
                    // Update the Power Query formula
                    string newFormula = "let Source = Csv.Document(File.Contents(\"C:\\\\NewFolder\\\\newfile.csv\"),[Delimiter=\",\", Columns=5, Encoding=1252, QuoteStyle=QuoteStyle.None]) in Source";
                    valueProp.SetValue(item, newFormula);
                    Console.WriteLine("Power Query formula updated successfully.");
                    itemFound = true;
                    break;
                }
            }

            if (!itemFound)
            {
                Console.WriteLine("PowerQueryFormulaItem with the specified name was not found.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
