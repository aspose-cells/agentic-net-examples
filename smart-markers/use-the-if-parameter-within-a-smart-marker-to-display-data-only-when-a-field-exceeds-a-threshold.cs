// Title: How to use an IF smart marker in Aspose.Cells (C#) to display scores only above a given threshold
// AI Prompts: Write C# code that creates a workbook, defines a smart‑marker range, and uses the IF function to show the Score column only when the value exceeds a specified limit. | Generate a reusable C# method that accepts a data collection and a numeric threshold, then applies WorkbookDesigner with an IF smart marker to conditionally display the field. | Adapt the example to output "Pass" or "Fail" instead of the raw score by using a configurable threshold in an Aspose.Cells IF smart marker.
// Common Searches: Aspose.Cells C# smart marker IF expression threshold example | how to conditionally hide Excel cell values with Aspose.Cells smart markers | process a named smart marker range with WorkbookDesigner in C# | display only high scores using smart markers in Aspose.Cells | C# Aspose.Cells conditional smart marker for numeric fields
// Tags: Aspose.Cells IF smart marker C# | WorkbookDesigner conditional display | named range smart marker processing | threshold-based smart marker | Excel export with conditional smart markers

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The sample creates a workbook, adds headers, inserts smart markers—including an IF expression that shows the Score only when it exceeds 80—binds a List<Person> as the data source, defines a named range for the markers, processes that range with WorkbookDesigner while preserving unrecognized markers, and saves the result to SmartMarkerIfOutput.xlsx.
class SmartMarkerIfDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add column headers
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["C1"].PutValue("High Score");

            // Insert smart markers (line‑by‑line mode)
            // &Name and &Score will be replaced directly
            // &IF(Score>80,Score,"") will display the score only when it exceeds 80
            sheet.Cells["A2"].PutValue("&=Name");
            sheet.Cells["B2"].PutValue("&=Score");
            sheet.Cells["C2"].PutValue("&=IF(Score>80,Score,\"\")");

            // Prepare sample data
            List<Person> data = new List<Person>
            {
                new Person { Name = "Alice",   Score = 75 },
                new Person { Name = "Bob",     Score = 92 },
                new Person { Name = "Charlie", Score = 68 }
            };

            // Set up the WorkbookDesigner and bind the data source
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource("Data", data);

            // Define the range that contains the smart markers and give it the required name
            AsposeRange smartMarkerRange = sheet.Cells.CreateRange("A2:C2");
            smartMarkerRange.Name = "_CellsSmartMarkers";

            // Process only the defined range (true = preserve unrecognized markers)
            designer.Process(smartMarkerRange, true);

            // Save the resulting workbook
            string outputPath = "SmartMarkerIfOutput.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Simple data class used as the data source
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Score { get; set; }
    }
}
