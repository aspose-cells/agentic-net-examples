using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        TriggerInvalidSheetName();
        TriggerInvalidFormula();
        TriggerErrorCellValue();
    }

    // -------------------------------------------------------------------------
    // Example: Invalid worksheet name (CellsExceptionType.SheetName)
    // -------------------------------------------------------------------------
    static void TriggerInvalidSheetName()
    {
        try
        {
            Workbook wb = new Workbook();
            wb.Worksheets[0].Name = "Invalid/Name";
            wb.Save("InvalidSheetName.xlsx");
        }
        catch (CellsException ex)
        {
            Console.WriteLine("Caught CellsException for sheet name:");
            Console.WriteLine($"  Code    : {ex.Code}");
            Console.WriteLine($"  Message : {ex.Message}");
        }
    }

    // -------------------------------------------------------------------------
    // Example: Invalid formula (CellsExceptionType.Formula)
    // -------------------------------------------------------------------------
    static void TriggerInvalidFormula()
    {
        try
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].Formula = "=1/0";
            wb.CalculateFormula();
        }
        catch (CellsException ex)
        {
            Console.WriteLine("Caught CellsException for formula error:");
            Console.WriteLine($"  Code    : {ex.Code}");
            Console.WriteLine($"  Message : {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected exception: {ex.Message}");
        }
    }

    // -------------------------------------------------------------------------
    // Example: Cell contains an error value
    // -------------------------------------------------------------------------
    static void TriggerErrorCellValue()
    {
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Cells["A1"].Formula = "=UNIQUE({1,2,3;1,2,3})";

        try
        {
            wb.CalculateFormula();
        }
        catch
        {
            // Ignore calculation errors for this demo
        }

        Cell cell = ws.Cells["A1"];
        if (cell.Type == CellValueType.IsError)
        {
            Console.WriteLine($"Cell A1 contains error (code {cell.IntValue}).");
        }
        else
        {
            Console.WriteLine("Cell A1 does not contain an error.");
        }
    }
}