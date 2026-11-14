// Title: Access the second worksheet in an Aspose.Cells workbook to verify formulas using C#
// Common Searches: Aspose.Cells C# get worksheet at index 1 for formula calculation | How to select the second sheet in a workbook with Aspose.Cells before recalculating formulas | C# example to load Excel file and access specific worksheet by position using Aspose.Cells | Retrieve second worksheet for formula verification in Aspose.Cells .NET | Using Aspose.Cells to recalculate formulas on a particular worksheet in C#
// Tags: worksheet index retrieval Aspose.Cells C# | second sheet access for formula evaluation Aspose.Cells | calculate formulas on selected worksheet Aspose.Cells | load workbook and select worksheet by position C# | Aspose.Cells formula verification on specific sheet

using Aspose.Cells;

// Load the workbook (replace with your actual file path or stream)
Workbook workbook = new Workbook("input.xlsx");

// Access the second worksheet (zero‑based index, so index 1 is the second sheet)
Worksheet secondWorksheet = workbook.Worksheets[1];

// At this point you can perform formula verification on `secondWorksheet`
// e.g., workbook.CalculateFormula(); // if you need to recalculate formulas.
