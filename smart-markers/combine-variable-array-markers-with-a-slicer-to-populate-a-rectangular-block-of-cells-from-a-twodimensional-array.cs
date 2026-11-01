// Title: Use Aspose.Cells smart markers with slicer syntax to fill a rectangular range from a 2D int array in C#
// Common Searches: how to apply slicer syntax in Aspose.Cells smart markers to extract a sub‑array | populate an Excel range with a portion of a two‑dimensional int array using WorkbookDesigner C# | Aspose.Cells variable array marker example with slice [1:3,2:4] in C# | fill cells A1:D3 from a 5x5 array using Aspose.Cells smart markers | set data source for smart marker from int[,] and process with Aspose.Cells
// Tags: smart markers slicer syntax Aspose.Cells | populate rectangular range from int[,] Aspose.Cells | WorkbookDesigner SetDataSource two dimensional array | variable array marker slice extraction | C# Aspose.Cells fill Excel range with subarray

// Create a new workbook
var workbook = new Aspose.Cells.Workbook();

// Access the first worksheet
var sheet = workbook.Worksheets[0];

// Place a variable array marker with a slicer in cell A1.
// This marker will extract rows 1 to 3 and columns 2 to 4 (zero‑based) from the data source named "Data".
sheet.Cells["A1"].PutValue("&=Data[1:3,2:4]");

// Prepare a two‑dimensional array (5x5) with sample values
int[,] data = new int[5, 5];
for (int i = 0; i < 5; i++)
{
    for (int j = 0; j < 5; j++)
    {
        data[i, j] = (i + 1) * 10 + (j + 1); // Example: 11,12,...,55
    }
}

// Use WorkbookDesigner to process smart markers
var designer = new Aspose.Cells.WorkbookDesigner(workbook);

// Set the data source for the marker; the name must match the marker ("Data")
designer.SetDataSource("Data", data);

// Process the smart markers – the slicer will populate a rectangular block starting at A1
designer.Process();

// Save the resulting workbook
workbook.Save("Output.xlsx");
