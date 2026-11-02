// Title: Retrieve a range of values returned by a custom array function after calculating formulas with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells get 2D array result from a custom Excel function after CalculateFormula | how to extract values from a cell that returns a range in Aspose.Cells | iterate over custom function returned range in a worksheet using Aspose.Cells .NET | retrieve object[,] output of a user‑defined function in Aspose.Cells | save evaluated workbook after processing custom array function results in C#
// Tags: custom function array result Aspose.Cells | extract 2d object array from evaluated cell | calculate formulas with user‑defined functions .NET | iterate over returned range values C# | retrieve range values from custom Excel function

// Load an existing workbook (replace with your file path)
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(@"C:\Path\To\YourWorkbook.xlsx");

// Force calculation of all formulas (including custom functions)
workbook.CalculateFormula();

// Get the cell that contains the custom function (adjust the address as needed)
Aspose.Cells.Worksheet worksheet = workbook.Worksheets[0];
Aspose.Cells.Cell targetCell = worksheet.Cells["A1"];   // e.g., =MyCustomArrayFunction()

// The custom function returns an array/range. After calculation Aspose.Cells stores it as a 2‑D object array.
object cellValue = targetCell.Value;

// Check if the result is an array and retrieve its values
if (cellValue is object[,] resultArray)
{
    int rows = resultArray.GetLength(0);
    int cols = resultArray.GetLength(1);

    // Example: iterate through the returned range and print each value
    for (int r = 0; r < rows; r++)
    {
        for (int c = 0; c < cols; c++)
        {
            // Each element can be cast to its actual type (e.g., double, string, etc.)
            object value = resultArray[r, c];
            System.Console.WriteLine($"Row {r + 1}, Column {c + 1}: {value}");
        }
    }
}
else
{
    // If the function does not return an array, handle the scalar result
    System.Console.WriteLine($"Scalar result: {cellValue}");
}

// (Optional) Save the workbook if any changes were made
workbook.Save(@"C:\Path\To\YourWorkbook_Evaluated.xlsx");
