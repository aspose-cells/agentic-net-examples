// Title: C# – Calculate Age with a Smart Marker Using DATEDIF in Aspose.Cells for .NET
// Description: This Aspose.Cells for .NET example creates a template workbook, inserts smart markers for Name, BirthDate and Age, and applies the Formula parameter with DATEDIF and TODAY to compute a person's age. The smart‑marker range is named _CellsSmartMarkers, a DataTable provides the source data, WorkbookDesigner processes the markers with CalculateFormula enabled, and the final file is saved as an XLSX workbook.
// Keywords: Aspose.Cells | smart markers | C# | .NET | calculate age | DATEDIF | TODAY function | WorkbookDesigner | template workbook | date functions | GitHub example
// Common Searches: Aspose.Cells calculate age smart marker | DATEDIF formula in Aspose.Cells | C# smart marker age example | How to use Formula parameter with date functions in Aspose.Cells | Smart markers date calculation .NET
// Developer Intent: Create a spreadsheet where the Age column is automatically derived from a BirthDate column using a smart‑marker formula.
// Use Cases: Employee roster that always shows current ages without storing age in the database | Customer list with dynamic age calculation for marketing segmentation | Birthday reminder sheet that updates ages each time the report is generated
// AI Prompts: Generate C# code with Aspose.Cells that adds a smart marker to compute age from a birthdate using DATEDIF and TODAY. | Explain how to enable formula evaluation for smart markers in WorkbookDesigner when performing date calculations. | Show the steps to define the _CellsSmartMarkers range for a smart marker that includes a Formula parameter.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerAgeDemo
{
    // This Aspose.Cells for .NET example creates a template workbook, inserts smart markers for Name, BirthDate and Age, and applies the Formula parameter with DATEDIF and TODAY to compute a person's age. The smart‑marker range is named _CellsSmartMarkers, a DataTable provides the source data, WorkbookDesigner processes the markers with CalculateFormula enabled, and the final file is saved as an XLSX workbook.
    class Program
    {
        static void Main()
        {
            // -------------------- Create template workbook --------------------
            Workbook template = new Workbook();
            Worksheet sheet = template.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("BirthDate");
            cells["C1"].PutValue("Age");

            // Smart markers
            // Name smart marker
            cells["A2"].PutValue("&=[People].[Name]");

            // BirthDate smart marker (will be filled with the actual date)
            cells["B2"].PutValue("&=[People].[BirthDate]");

            // Age smart marker using Formula parameter with DATEDIF to calculate years between birthdate and today
            // The smart marker syntax: &=[People].[BirthDate] with Formula="=DATEDIF(&=BirthDate,TODAY(),\"Y\")"
            cells["C2"].PutValue("&=[People].[BirthDate] with Formula=\"=DATEDIF(&=BirthDate,TODAY(),\\\"Y\\\")\"");

            // Define the range that contains smart markers (required for processing)
            // The name "_CellsSmartMarkers" is recognized by WorkbookDesigner
            sheet.Cells.CreateRange("A2:C2").Name = "_CellsSmartMarkers";

            // -------------------- Prepare data source --------------------
            DataTable people = new DataTable("People");
            people.Columns.Add("Name", typeof(string));
            people.Columns.Add("BirthDate", typeof(DateTime));

            // Sample data
            people.Rows.Add("John Doe", new DateTime(1990, 5, 15));
            people.Rows.Add("Jane Smith", new DateTime(1985, 12, 3));
            people.Rows.Add("Bob Johnson", new DateTime(2000, 8, 22));

            // -------------------- Process smart markers --------------------
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = template;               // Load the template workbook
            designer.SetDataSource(people);             // Bind the DataTable as data source
            designer.CalculateFormula = true;           // Ensure formulas (age calculation) are evaluated
            designer.Process();                         // Populate smart markers

            // -------------------- Save result --------------------
            // Save to a memory stream (demonstrates using a stream) and then to a file
            using (MemoryStream ms = new MemoryStream())
            {
                designer.Workbook.Save(ms, SaveFormat.Xlsx);
                File.WriteAllBytes("PeopleWithAge.xlsx", ms.ToArray());
            }

            Console.WriteLine("Workbook generated: PeopleWithAge.xlsx");
        }
    }
}
