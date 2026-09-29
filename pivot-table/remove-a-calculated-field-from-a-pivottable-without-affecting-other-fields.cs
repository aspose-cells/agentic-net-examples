// Title: Delete a calculated field from a PivotTable in an existing Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that removes a calculated field named 'MyCalculatedField' from the first PivotTable in a workbook while leaving other fields untouched. | Show how to check whether a calculated field exists in a PivotTable before deleting it, using dynamic typing to support different Aspose.Cells versions. | Provide a .NET example that accesses the CalculatedFields collection dynamically, removes the target field, and safely saves the modified Excel file.
// Common Searches: asp.net remove calculated field from pivot table using Aspose.Cells | c# code to delete specific calculated field in Excel pivot table | how to verify calculated field existence before removal with Aspose.Cells | dynamic access to CalculatedFields collection Aspose.Cells version compatibility | remove pivot table calculated field without affecting other fields in .NET
// Tags: Aspose.Cells PivotTable calculated field deletion | C# remove pivot calculated field | Aspose.Cells CalculatedFields dynamic handling | Excel workbook pivot field management .NET | version‑compatible Aspose.Cells pivot operations

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable

// The example loads an Excel workbook, locates the first PivotTable, checks for a calculated field named 'MyCalculatedField', removes it using dynamic access to the CalculatedFields collection (compatible with multiple Aspose.Cells versions), and saves the workbook while handling possible errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the existing workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one pivot table
        if (worksheet.PivotTables == null || worksheet.PivotTables.Count == 0)
        {
            Console.WriteLine("No pivot tables found on the worksheet.");
            return;
        }

        // Get the first pivot table (adjust index or name as required)
        PivotTable pivotTable = worksheet.PivotTables[0];

        // Name of the calculated field to be removed
        string calculatedFieldName = "MyCalculatedField";

        // Attempt to remove the calculated field using dynamic to stay compatible with different library versions
        try
        {
            dynamic dynPivot = pivotTable;
            // Try to access the CalculatedFields collection; if it doesn't exist, an exception will be thrown
            dynamic calcFields = dynPivot.CalculatedFields;

            bool fieldExists = false;
            try
            {
                // Some versions expose Contains(string), others may require iteration
                fieldExists = calcFields.Contains(calculatedFieldName);
            }
            catch
            {
                // Fallback: iterate through the collection to find the field by name
                foreach (var cf in calcFields)
                {
                    if (cf.Name == calculatedFieldName)
                    {
                        fieldExists = true;
                        break;
                    }
                }
            }

            if (fieldExists)
            {
                // Remove the calculated field
                calcFields.Remove(calculatedFieldName);
                Console.WriteLine($"Calculated field '{calculatedFieldName}' removed successfully.");
            }
            else
            {
                Console.WriteLine($"Calculated field '{calculatedFieldName}' does not exist.");
            }
        }
        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
        {
            // The CalculatedFields property is not available in this version of Aspose.Cells
            Console.WriteLine("Calculated fields are not supported by the current Aspose.Cells version.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during removal
            Console.WriteLine($"Unable to remove calculated field '{calculatedFieldName}': {ex.Message}");
        }

        // Save the modified workbook
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
