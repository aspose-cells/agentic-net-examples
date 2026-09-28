// Title: Add a master smart marker row in Aspose.Cells (C#) to automatically repeat rows for each item in a List<T>
// AI Prompts: Write C# code that inserts a master smart marker in an Aspose.Cells worksheet, defines the _CellsSmartMarkers range, binds a List<MasterRecord> as the data source, and processes the template to generate repeated rows. | Show how to use WorkbookDesigner in Aspose.Cells to bind a collection of objects to a master smart marker and export the result to an Excel file.
// Common Searches: Aspose.Cells C# master smart marker repeat rows for each list element | Define _CellsSmartMarkers range for master smart markers in Aspose.Cells | Binding a List of custom objects to a master smart marker in Aspose.Cells C# | Generate Excel rows from a collection using Aspose.Cells smart markers without LineByLine | Create master smart marker template row in Aspose.Cells workbook programmatically
// Tags: row duplication using master smart marker Aspose.Cells | WorkbookDesigner bind collection to smart marker | smart marker range definition _CellsSmartMarkers | process smart markers with default linebyline | repeat Excel rows programmatically Aspose.Cells C#

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // Sample master record class
    // The example demonstrates how to create a new workbook, add header cells, place master smart markers in a template row, define the _CellsSmartMarkers range, bind a List<MasterRecord> as the data source named 'MasterData', process the markers with WorkbookDesigner to repeat the row for each record, and save the resulting workbook as an .xlsx file.
    public class MasterRecord
    {
        public string? Name { get; set; }   // Made nullable to satisfy non‑nullable warning
        public int Age { get; set; }
    }

    public class MasterSmartMarkerExample
    {
        public static void Run()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Set up header cells
                sheet.Cells["A1"].PutValue("Name");
                sheet.Cells["B1"].PutValue("Age");

                // 3. Place master smart markers in the template row (these will be repeated)
                //    The syntax "&=MasterData.Property" tells the designer to repeat this row
                //    for each item in the data source named "MasterData".
                sheet.Cells["A2"].PutValue("&=MasterData.Name");
                sheet.Cells["B2"].PutValue("&=MasterData.Age");

                // 4. Define the range that contains the smart markers.
                //    Naming the range "_CellsSmartMarkers" enables processing when LineByLine is false.
                Aspose.Cells.Range smartMarkerRange = sheet.Cells.CreateRange("A2:B2");
                smartMarkerRange.Name = "_CellsSmartMarkers";

                // 5. Prepare sample data source (a list of master records)
                List<MasterRecord> masterData = new List<MasterRecord>
                {
                    new MasterRecord { Name = "John Doe", Age = 30 },
                    new MasterRecord { Name = "Jane Smith", Age = 28 },
                    new MasterRecord { Name = "Bob Johnson", Age = 45 }
                };

                // 6. Initialize WorkbookDesigner, bind the workbook and data source
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                    // LineByLine defaults to false; range smart markers will be used.
                };
                designer.SetDataSource("MasterData", masterData);

                // 7. Process the smart markers – rows will be repeated for each master record
                designer.Process();

                // 8. Save the resulting workbook
                string outputPath = "MasterSmartMarkerOutput.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during smart marker processing: {ex.Message}");
            }
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            MasterSmartMarkerExample.Run();
            Console.WriteLine("Workbook with master smart marker created successfully.");
        }
    }
}
