using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class WorksheetScopedNamedRangeDemo
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the default first worksheet and rename it
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";

        // Add a second worksheet where the scoped name will be defined
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

        // Populate some sample data in Sheet2
        sheet2.Cells["A1"].PutValue("Alpha");
        sheet2.Cells["A2"].PutValue("Beta");
        sheet2.Cells["A3"].PutValue("Gamma");

        // ------------------------------------------------------------
        // Create a worksheet‑scoped named range on Sheet2
        // ------------------------------------------------------------
        // Add a new name to the workbook's name collection (initially global)
        int nameIndex = workbook.Worksheets.Names.Add("LocalRange");
        Name localName = workbook.Worksheets.Names[nameIndex];

        // Set the scope to Sheet2 (one‑based sheet index, so 2 for the second sheet)
        localName.SheetIndex = 2;

        // Define the range that the name refers to (absolute reference)
        localName.RefersTo = "=Sheet2!$A$1:$A$3";

        // ------------------------------------------------------------
        // Retrieve and manipulate the named range using the Name object
        // ------------------------------------------------------------
        AsposeRange rangeViaName = localName.GetRange(); // Returns the range on Sheet2

        // Example manipulation: append "_Modified" to each cell's text
        foreach (Cell cell in rangeViaName)
        {
            string current = cell.StringValue;
            cell.PutValue(current + "_Modified");
        }

        // ------------------------------------------------------------
        // Alternative retrieval using WorksheetCollection.GetRangeByName
        // ------------------------------------------------------------
        // Parameters: (range name, current sheet index (zero‑based), include tables)
        AsposeRange rangeViaCollection = workbook.Worksheets.GetRangeByName("LocalRange", 1, true);

        // Demonstrate that the retrieved range belongs to Sheet2
        Console.WriteLine("Range belongs to worksheet: " + rangeViaCollection.Worksheet.Name);

        // ------------------------------------------------------------
        // Save the modified workbook
        // ------------------------------------------------------------
        workbook.Save("output.xlsx");
    }
}