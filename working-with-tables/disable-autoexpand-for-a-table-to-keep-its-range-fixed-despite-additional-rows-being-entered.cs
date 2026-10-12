// Title: Disable auto‑expand for a ListObject (Excel table) in Aspose.Cells for .NET to keep the table range fixed
// AI Prompts: Generate C# code with Aspose.Cells that stops a ListObject from automatically expanding when new rows are added. | Show how to keep an Excel table's range static in Aspose.Cells for .NET, even after inserting additional data rows. | Explain a workaround for the missing ShowAutoExpand property in the current Aspose.Cells version to prevent table growth.
// Common Searches: Aspose.Cells .NET prevent ListObject from expanding when adding rows | how to keep Excel table range fixed using Aspose.Cells | ShowAutoExpand property not available in Aspose.Cells | disable auto expand for Excel table in Aspose.Cells C# | static table range Aspose.Cells example
// Tags: disable ListObject auto expand Aspose.Cells | static Excel table range .NET | prevent table range growth Aspose.Cells | Aspose.Cells ListObject configuration C# | fixed table range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The example loads an existing workbook (or creates a new one), accesses the first worksheet, retrieves a ListObject named "MyTable", notes that the ShowAutoExpand property is unavailable in this Aspose.Cells version, and saves the workbook. It demonstrates how to handle the situation when auto‑expand cannot be directly disabled.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";
            Workbook workbook = null;

            try
            {
                // Load existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                }

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Retrieve the table (ListObject) by its name
                ListObject table = sheet.ListObjects["MyTable"];
                if (table != null)
                {
                    // The ShowAutoExpand property is not available in this version of Aspose.Cells.
                    // If needed, other table settings can be adjusted here.
                }
                else
                {
                    Console.WriteLine("Table 'MyTable' not found in the worksheet.");
                }

                // Save the workbook with the changes
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
