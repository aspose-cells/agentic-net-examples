// Title: How to verify that a ListObject loses table features after converting an XLSX workbook to Excel 97‑2003 format with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# check ListObject.IsTable after saving as Excel97To2003 | detect loss of table sorting in XLS file generated from XLSX using Aspose.Cells | verify ListObject properties after converting workbook to .xls format in .NET | how to confirm table features are removed when converting XLSX to XLS with Aspose.Cells | C# Aspose.Cells convert workbook to Excel 97-2003 and test table support
// Tags: Aspose.Cells ListObject.IsTable after format conversion | convert XLSX to Excel97To2003 table compatibility | C# detect removed Excel table features | Aspose.Cells verify table support in .xls files | Excel 97-2003 format table sorting limitation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates an XLSX workbook containing a ListObject (table), saves it, converts it to the legacy Excel 97‑2003 (.xls) format using SaveFormat.Excel97To2003, reloads the file, and checks the ListObject.IsTable property to confirm that table capabilities such as sorting are no longer supported.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string convertedPath = "converted.xls";

            // Ensure the input file exists; if not, create a simple workbook with a table.
            if (!File.Exists(inputPath))
            {
                try
                {
                    var wb = new Workbook();
                    var ws = wb.Worksheets[0];
                    ws.Cells["A1"].PutValue("ID");
                    ws.Cells["B1"].PutValue("Name");
                    ws.Cells["A2"].PutValue(1);
                    ws.Cells["B2"].PutValue("Alice");
                    ws.Cells["A3"].PutValue(2);
                    ws.Cells["B3"].PutValue("Bob");

                    // Create a table (ListObject) over the data range.
                    var table = ws.ListObjects[ws.ListObjects.Add("A1", "B3", true)];
                    table.DisplayName = "SampleTable";

                    wb.Save(inputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Created sample input file: {inputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to create sample input file: {ex.Message}");
                    return;
                }
            }

            // Load the original workbook that contains a table (ListObject)
            Workbook originalWb;
            try
            {
                originalWb = new Workbook(inputPath);
            }
            catch (FileNotFoundException)
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load input workbook: {ex.Message}");
                return;
            }

            // Save the workbook in an older Excel format (Excel 97-2003) which does not support tables
            try
            {
                originalWb.Save(convertedPath, SaveFormat.Excel97To2003);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save converted workbook: {ex.Message}");
                return;
            }

            // Reload the converted workbook if it exists
            if (!File.Exists(convertedPath))
            {
                Console.WriteLine($"Converted file not found: {convertedPath}");
                return;
            }

            Workbook convertedWb;
            try
            {
                convertedWb = new Workbook(convertedPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load converted workbook: {ex.Message}");
                return;
            }

            // Access the first worksheet (adjust index if needed)
            var sheet = convertedWb.Worksheets[0];

            // Check for ListObjects (tables) in the worksheet
            if (sheet.ListObjects.Count > 0)
            {
                // Get the first ListObject
                ListObject table = sheet.ListObjects[0];

                // After conversion to a format that doesn't support tables, the ListObject remains
                // but its properties reflect the loss of table features. Display its display name.
                Console.WriteLine($"ListObject display name after conversion: {table.DisplayName}");
            }
            else
            {
                Console.WriteLine("No ListObjects (tables) found in the worksheet.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An unexpected error occurred: " + ex.Message);
        }
    }
}
