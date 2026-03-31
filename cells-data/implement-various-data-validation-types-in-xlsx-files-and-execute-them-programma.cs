using System;
using Aspose.Cells;

class DataValidationDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Add different types of data validation
        AddWholeNumberValidation(ws, "A1", 1, 100);
        AddDecimalValidation(ws, "B1", 0.0, 10.0);
        AddListValidation(ws, "C1", "Red,Green,Blue");
        AddDateValidation(ws, "D1", "2023-01-01", "2023-12-31");
        AddTimeValidation(ws, "E1", "09:00", "18:00");
        AddTextLengthValidation(ws, "F1", 5, 10);
        AddCustomValidation(ws, "G1", "=LEN(A1)>3");

        // Save the workbook (disable Excel restriction checking for safety)
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.CheckExcelRestriction = false;
        wb.Save("DataValidations.xlsx", saveOptions);

        // Load the workbook with data‑validation checking enabled
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.CheckDataValid = true; // enforce validation on load
        Workbook loadedWb = new Workbook("DataValidations.xlsx", loadOptions);
        Worksheet loadedWs = loadedWb.Worksheets[0];

        // Attempt to put invalid values and catch validation exceptions
        TryPutValue(loadedWs, "A1", 200);          // WholeNumber out of range
        TryPutValue(loadedWs, "B1", 20.5);         // Decimal out of range
        TryPutValue(loadedWs, "C1", "Yellow");    // List value not allowed
        TryPutValue(loadedWs, "D1", "2024-01-01"); // Date out of range
        TryPutValue(loadedWs, "E1", "08:00");      // Time out of range
        TryPutValue(loadedWs, "F1", "abc");        // Text length too short
        TryPutValue(loadedWs, "G1", 5);            // Custom formula fails

        // Display validation information for a couple of cells
        PrintValidationInfo(loadedWs, "A1");
        PrintValidationInfo(loadedWs, "C1");
    }

    static void AddWholeNumberValidation(Worksheet ws, string cellName, int min, int max)
    {
        Validation v = ws.Validations[ws.Validations.Add()];
        v.Type = ValidationType.WholeNumber;
        v.Operator = OperatorType.Between;
        v.Formula1 = min.ToString();
        v.Formula2 = max.ToString();
        v.InputMessage = $"Enter whole number between {min} and {max}.";
        v.ErrorMessage = $"Value must be between {min} and {max}.";
        v.ShowInput = true;
        v.ShowError = true;
        v.AddArea(CellArea.CreateCellArea(cellName, cellName));
    }

    static void AddDecimalValidation(Worksheet ws, string cellName, double min, double max)
    {
        Validation v = ws.Validations[ws.Validations.Add()];
        v.Type = ValidationType.Decimal;
        v.Operator = OperatorType.Between;
        v.Formula1 = min.ToString();
        v.Formula2 = max.ToString();
        v.AddArea(CellArea.CreateCellArea(cellName, cellName));
    }

    static void AddListValidation(Worksheet ws, string cellName, string list)
    {
        Validation v = ws.Validations[ws.Validations.Add()];
        v.Type = ValidationType.List;
        v.InCellDropDown = true;
        v.Formula1 = list; // comma‑separated list
        v.AddArea(CellArea.CreateCellArea(cellName, cellName));
    }

    static void AddDateValidation(Worksheet ws, string cellName, string startDate, string endDate)
    {
        Validation v = ws.Validations[ws.Validations.Add()];
        v.Type = ValidationType.Date;
        v.Operator = OperatorType.Between;
        v.Formula1 = startDate;
        v.Formula2 = endDate;
        v.AddArea(CellArea.CreateCellArea(cellName, cellName));
    }

    static void AddTimeValidation(Worksheet ws, string cellName, string startTime, string endTime)
    {
        Validation v = ws.Validations[ws.Validations.Add()];
        v.Type = ValidationType.Time;
        v.Operator = OperatorType.Between;
        v.Formula1 = startTime;
        v.Formula2 = endTime;
        v.AddArea(CellArea.CreateCellArea(cellName, cellName));
    }

    static void AddTextLengthValidation(Worksheet ws, string cellName, int minLen, int maxLen)
    {
        Validation v = ws.Validations[ws.Validations.Add()];
        v.Type = ValidationType.TextLength;
        v.Operator = OperatorType.Between;
        v.Formula1 = minLen.ToString();
        v.Formula2 = maxLen.ToString();
        v.AddArea(CellArea.CreateCellArea(cellName, cellName));
    }

    static void AddCustomValidation(Worksheet ws, string cellName, string formula)
    {
        Validation v = ws.Validations[ws.Validations.Add()];
        v.Type = ValidationType.Custom;
        v.Formula1 = formula;
        v.AddArea(CellArea.CreateCellArea(cellName, cellName));
    }

    static void TryPutValue(Worksheet ws, string cellName, object value)
    {
        try
        {
            ws.Cells[cellName].PutValue(value);
            Console.WriteLine($"Successfully set {cellName} = {value}");
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.DataValidation)
        {
            Console.WriteLine($"Data validation error at {cellName}: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error at {cellName}: {ex.Message}");
        }
    }

    static void PrintValidationInfo(Worksheet ws, string cellName)
    {
        int row = ws.Cells[cellName].Row;
        int col = ws.Cells[cellName].Column;
        Validation v = ws.Validations.GetValidationInCell(row, col);
        if (v != null)
        {
            Console.WriteLine($"Cell {cellName} validation type: {v.Type}");
            Console.WriteLine($"Formula1: {v.Formula1}");
            Console.WriteLine($"Formula2: {v.Formula2}");
        }
        else
        {
            Console.WriteLine($"No validation found for cell {cellName}");
        }
    }
}