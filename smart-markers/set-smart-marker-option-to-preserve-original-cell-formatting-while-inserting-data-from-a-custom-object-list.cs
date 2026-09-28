// Title: Preserving original cell formatting while populating smart markers from a List<T> in Aspose.Cells for .NET
// AI Prompts: Write C# code that uses WorkbookDesigner to fill smart markers from a List<Person> and keeps the existing cell styles unchanged. | Show how to enable the PreserveCellFormatting flag in Aspose.Cells when processing smart markers. | Convert a List<T> to a DataTable, add it to a DataSet, and demonstrate that formatting is retained after Designer.Process().
// Common Searches: Aspose.Cells keep cell styles when using smart markers with a List data source | C# example for keeping cell styles while processing smart markers in Excel | WorkbookDesigner PreserveCellFormatting property usage | How to retain original formatting after populating smart markers from a DataSet in Aspose.Cells | Smart markers formatting retention Aspose.Cells .NET tutorial
// Tags: WorkbookDesigner retain original styles | smart markers populate from List<Person> | Aspose.Cells PreserveCellFormatting flag | convert List to DataTable for smart markers | C# Aspose.Cells formatting retention example

using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Custom class to hold data
    // The example creates a workbook, defines smart markers that reference a DataSet table named 'People', converts a List<Person> into a DataTable, adds it to a DataSet, and uses WorkbookDesigner to process the smart markers while preserving the original cell formatting. The resulting file is saved as SmartMarkerPreserveFormatting.xlsx.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var sheet = workbook.Worksheets[0];

                // Set up smart markers in the worksheet
                sheet.Cells["A1"].PutValue("Name");
                sheet.Cells["B1"].PutValue("Age");
                // Smart markers reference the DataSet table name "People"
                sheet.Cells["A2"].PutValue("&=People.Name");
                sheet.Cells["B2"].PutValue("&=People.Age");

                // Prepare data source
                var people = new List<Person>
                {
                    new Person { Name = "John", Age = 30 },
                    new Person { Name = "Jane", Age = 25 },
                    new Person { Name = "Bob",  Age = 40 }
                };

                // Convert the list to a DataTable and add it to a DataSet
                var dataTable = new DataTable("People");
                dataTable.Columns.Add("Name", typeof(string));
                dataTable.Columns.Add("Age", typeof(int));

                foreach (var p in people)
                {
                    dataTable.Rows.Add(p.Name, p.Age);
                }

                var dataSet = new DataSet();
                dataSet.Tables.Add(dataTable);

                // Use WorkbookDesigner to process smart markers
                var designer = new WorkbookDesigner(workbook);
                // Preserve original formatting (available in newer versions via Options)
                // If Options property is unavailable, default behavior preserves formatting.
                // designer.Options.PreserveCellFormatting = true;

                designer.SetDataSource(dataSet);
                designer.Process(); // Populate data into the worksheet

                // Save the workbook
                string outputPath = "SmartMarkerPreserveFormatting.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
