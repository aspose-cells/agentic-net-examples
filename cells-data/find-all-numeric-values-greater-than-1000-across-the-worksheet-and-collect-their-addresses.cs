// Title: Extract cell addresses of numeric values greater than 1000 using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# find cells with numbers larger than 1000 and get their addresses | how to list Excel cell references where value exceeds 1000 with Aspose.Cells | C# iterate worksheet cells and collect addresses of values over 1000 using Aspose.Cells | filter numeric cells above 1000 in Aspose.Cells and retrieve cell names | retrieve high-value cell locations from Excel file using Aspose.Cells .NET
// Tags: filter numeric cells above 1000 Aspose.Cells | extract cell addresses C# Aspose.Cells | iterate worksheet used range Aspose.Cells | Excel numeric threshold detection .NET | collect high-value cell references Aspose.Cells

// Load an existing workbook (replace with your file path)
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("input.xlsx");

// Get the first worksheet (or iterate through all worksheets if needed)
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Collection to hold the addresses of cells with numeric values > 1000
System.Collections.Generic.List<string> addresses = new System.Collections.Generic.List<string>();

// Iterate through all used cells in the worksheet
foreach (Aspose.Cells.Cell cell in sheet.Cells)
{
    // Check if the cell contains a numeric value
    if (cell.Type == Aspose.Cells.CellValueType.IsNumeric)
    {
        // Compare the numeric value with 1000
        if (cell.DoubleValue > 1000)
        {
            // Add the cell address (e.g., "B5") to the list
            addresses.Add(cell.Name);
        }
    }
}

// Example: output the collected addresses to the console
foreach (string addr in addresses)
{
    System.Console.WriteLine(addr);
}

// Save the workbook if any modifications were made (optional)
workbook.Save("output.xlsx");
