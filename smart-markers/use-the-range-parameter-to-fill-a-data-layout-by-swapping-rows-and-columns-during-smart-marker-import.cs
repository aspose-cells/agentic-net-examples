// Title: C# Example: Process a Smart‑Marker Range and Transpose Data with Aspose.Cells
// Description: Shows how to name a smart‑marker range, import a List<Person> via WorkbookDesigner.Process, then transpose the filled range to swap rows and columns and save the workbook.
// Keywords: Aspose.Cells | smart markers | WorkbookDesigner | Process range | Range.Transpose | C# | .NET | named range | data transpose | swap rows and columns | example | sample code
// Common Searches: Aspose.Cells process specific smart marker range | transpose smart marker output C# | swap rows and columns after smart marker import | Range.Transpose Aspose.Cells example | how to use WorkbookDesigner with a named range
// Developer Intent: The developer wants to import data only from a defined smart‑marker range and then rotate the resulting table so that rows become columns and vice‑versa.
// Use Cases: Import a collection of objects into a worksheet while leaving other sheet content untouched. | Convert a vertical list of names and scores into a horizontal layout after smart‑marker processing. | Create reusable templates that populate data in a specific area and then re‑orient the data for reporting.
// AI Prompts: Write C# code that defines a named smart‑marker range, processes it with WorkbookDesigner, and transposes the resulting cells. | Explain the effect of Range.Transpose on a smart‑marker output range in Aspose.Cells. | Provide step‑by‑step instructions to set up a smart‑marker range, import a List<T>, and swap rows and columns using Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Shows how to name a smart‑marker range, import a List<Person> via WorkbookDesigner.Process, then transpose the filled range to swap rows and columns and save the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Set up headers
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Score");

            // Place smart markers in a range that will be processed
            cells["A2"].PutValue("&=$People.Name");
            cells["B2"].PutValue("&=$People.Score");

            // Define the range that contains the smart markers and give it the special name
            AsposeRange smartMarkerRange = cells.CreateRange("A2:B2");
            smartMarkerRange.Name = "_CellsSmartMarkers";

            // Prepare a data source
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice",   Score = 85 },
                new Person { Name = "Bob",     Score = 92 },
                new Person { Name = "Charlie", Score = 78 }
            };

            // Configure the WorkbookDesigner with the data source
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource("People", people);

            // Process only the defined smart‑marker range (true = preserve unrecognized markers)
            designer.Process(smartMarkerRange, true);

            // After processing, the data occupies A2:B4. Transpose this range to swap rows and columns.
            AsposeRange dataRange = cells.CreateRange("A2:B4");
            dataRange.Transpose();

            // Save the final workbook
            workbook.Save("SmartMarkerTransposeOutput.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Simple POCO class used as the data source
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Score { get; set; }
    }
}
