// Title: Import a Sensor Readings Matrix with ImportArray and Add an Average Summary Row – Aspose.Cells for .NET (C#)
// Description: Creates a workbook, imports a double[,] sensor matrix horizontally using Cells.ImportArray, appends a labeled row, inserts AVERAGE formulas for each column, evaluates them with workbook.CalculateFormula, and saves the file as an Excel report.
// Keywords: Aspose.Cells | ImportArray | C# | .NET | sensor matrix import | Excel average row | column averages formula | workbook.CalculateFormula | CellsHelper.ColumnIndexToName | IoT data export | Excel automation
// Common Searches: Aspose.Cells ImportArray double array example | add summary row with averages in Aspose.Cells | calculate column averages programmatically C# | import 2D numeric matrix into Excel using Aspose.Cells | evaluate formulas after ImportArray Aspose.Cells
// Developer Intent: Load a numeric sensor matrix into a worksheet with ImportArray and automatically generate a row that shows column‑wise averages.
// Use Cases: Generate Excel reports from IoT sensor arrays with a built‑in average row for quick analysis. | Build a reusable routine that accepts any sized numeric matrix and appends a dynamic summary row using formulas. | Automate data logging pipelines that write sensor measurements to Excel and provide statistical summaries without manual editing.
// AI Prompts: Extend the code to compute row averages and place them in a new column beside the data. | Format the average cells with bold text and two decimal places after formula evaluation. | Show how to import the matrix vertically (column‑wise) using ImportArray and adjust the summary row accordingly.

using Aspose.Cells;
using System;

// Creates a workbook, imports a double[,] sensor matrix horizontally using Cells.ImportArray, appends a labeled row, inserts AVERAGE formulas for each column, evaluates them with workbook.CalculateFormula, and saves the file as an Excel report.
class SensorReadingsExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Sample sensor readings matrix (rows = sensors, columns = measurements)
        double[,] sensorReadings = new double[,]
        {
            { 10.5, 12.3, 11.0 },
            { 9.8, 13.2, 10.5 },
            { 11.4, 12.8, 12.0 }
        };

        int rowCount = sensorReadings.GetLength(0);
        int colCount = sensorReadings.GetLength(1);

        // Import each row horizontally using ImportArray(double[], firstRow, firstColumn, isVertical)
        for (int r = 0; r < rowCount; r++)
        {
            double[] rowData = new double[colCount];
            for (int c = 0; c < colCount; c++)
                rowData[c] = sensorReadings[r, c];

            // false => import horizontally
            cells.ImportArray(rowData, r, 0, false);
        }

        // Add a summary row after the data to hold column averages
        int summaryRowIndex = rowCount; // zero‑based index

        // Label for the summary row
        cells[summaryRowIndex, 0].PutValue("Average");

        // Insert AVERAGE formulas for each column
        for (int c = 0; c < colCount; c++)
        {
            // Convert column index to Excel column letter (A, B, C, ...)
            string colLetter = CellsHelper.ColumnIndexToName(c);
            // Build formula like AVERAGE(A1:A3)
            string formula = $"AVERAGE({colLetter}1:{colLetter}{rowCount})";
            cells[summaryRowIndex, c].Formula = formula;
        }

        // Evaluate formulas so that the average values are written to the cells
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("SensorReadingsWithAverages.xlsx");
    }
}
