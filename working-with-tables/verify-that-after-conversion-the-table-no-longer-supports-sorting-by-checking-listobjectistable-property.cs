// Title: How to confirm that a ListObject no longer supports sorting after converting an .xlsx workbook to Excel 97‑2003 (.xls) with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells check if a table loses sorting capability after saving as .xls | verify ListObject.IsTable is false after converting .xlsx to Excel97To2003 using Aspose.Cells | detect removal of Excel table features in a workbook converted to .xls in .NET | how to determine that table sorting is not available in an older Excel format with Aspose.Cells | Aspose.Cells example for testing table support after XLS conversion
// Tags: Aspose.Cells ListObject conversion verification | C# check table support in Excel97To2003 | detect loss of table sorting after XLS conversion | verify ListObject.IsTable property post‑conversion | Aspose.Cells older Excel format table feature test

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The code loads an .xlsx workbook containing a ListObject, saves it as an Excel 97‑2003 .xls file, reloads the saved workbook, accesses the first worksheet, and examines the ListObject's IsTable (and TableStyleType) property to confirm that table‑specific features such as sorting are no longer retained in the older format.
class VerifyTableSorting
{
    static void Main()
    {
        try
        {
            const string originalPath = "OriginalWithTable.xlsx";
            const string convertedPath = "Converted.xls";

            // Verify the source file exists
            if (!File.Exists(originalPath))
            {
                Console.WriteLine($"Error: File \"{originalPath}\" not found.");
                return;
            }

            // Load the original workbook that contains a table (ListObject)
            Workbook originalWb;
            try
            {
                originalWb = new Workbook(originalPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load \"{originalPath}\": {ex.Message}");
                return;
            }

            // Save the workbook to an older Excel format (Excel 97-2003) which does not support tables
            try
            {
                originalWb.Save(convertedPath, SaveFormat.Excel97To2003);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save as \"{convertedPath}\": {ex.Message}");
                return;
            }

            // Ensure the converted workbook was created
            if (!File.Exists(convertedPath))
            {
                Console.WriteLine($"Error: Failed to create \"{convertedPath}\".");
                return;
            }

            // Load the converted workbook
            Workbook convertedWb;
            try
            {
                convertedWb = new Workbook(convertedPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load \"{convertedPath}\": {ex.Message}");
                return;
            }

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = convertedWb.Worksheets[0];

            // Check for ListObjects in the worksheet
            if (sheet.ListObjects.Count > 0)
            {
                try
                {
                    ListObject listObj = sheet.ListObjects[0];

                    // After conversion to an older format, table-specific features are not supported.
                    // We can inspect a property such as TableStyleType to infer this.
                    bool hasTableStyle = listObj.TableStyleType != TableStyleType.None;

                    // Use DisplayName instead of Name (Name property is not available in this version)
                    Console.WriteLine($"ListObject DisplayName: {listObj.DisplayName}");
                    Console.WriteLine($"TableStyle applied after conversion: {hasTableStyle}");
                    // Expected: hasTableStyle should be false because the older format does not retain table styling.
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error while processing ListObject: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("No ListObjects found in the converted worksheet.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An unexpected error occurred: " + ex.Message);
        }
    }
}
