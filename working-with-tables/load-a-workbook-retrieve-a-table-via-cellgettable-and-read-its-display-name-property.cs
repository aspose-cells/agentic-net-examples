// Title: Read the DisplayName of an Excel ListObject (table) with Aspose.Cells for .NET in C#
// Common Searches: Aspose.Cells C# example to get Excel table display name | How to retrieve ListObject DisplayName property using Aspose.Cells | Read Excel table name from .xlsx file with Aspose.Cells for .NET | C# code to access first worksheet table and read its DisplayName | Aspose.Cells load workbook and extract table metadata in C#
// Tags: Aspose.Cells read ListObject display name | C# retrieve Excel table name Aspose.Cells | load workbook access first ListObject .NET | extract table metadata Aspose.Cells | Excel ListObject properties C#

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The sample checks for an 'input.xlsx' file, creates a workbook with a two‑column table named 'SampleTable' if needed, saves it, then loads the workbook, accesses the first worksheet, obtains the first ListObject, reads its DisplayName property, and prints the table's display name to the console.
class Program
{
    static void Main()
    {
        const string filePath = "input.xlsx";

        try
        {
            // Ensure the input file exists; create a sample workbook with a table if it does not.
            if (!File.Exists(filePath))
            {
                try
                {
                    Workbook sampleWb = new Workbook();
                    Worksheet sampleWs = sampleWb.Worksheets[0];

                    // Populate sample data
                    sampleWs.Cells["A1"].PutValue("Header1");
                    sampleWs.Cells["B1"].PutValue("Header2");
                    sampleWs.Cells["A2"].PutValue("Data1");
                    sampleWs.Cells["B2"].PutValue("Data2");

                    // Add a ListObject (table) covering the data range (A1:B2)
                    int firstRow = 0;          // zero‑based index for row 1
                    int firstColumn = 0;       // zero‑based index for column A
                    int totalRows = 2;         // includes header row
                    int totalColumns = 2;
                    int listObjectIndex = sampleWs.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
                    ListObject sampleListObj = sampleWs.ListObjects[listObjectIndex];
                    sampleListObj.DisplayName = "SampleTable";

                    sampleWb.Save(filePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Failed to create sample workbook: " + ex.Message);
                    return;
                }
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the first ListObject (table) in the worksheet, if any
            ListObject tableObj = worksheet.ListObjects.Count > 0 ? worksheet.ListObjects[0] : null;

            // Read the display name of the table
            string displayName = tableObj?.DisplayName ?? "No table found";

            // Output the display name
            Console.WriteLine("Table Display Name: " + displayName);
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
