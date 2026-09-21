// Title: Apply German culture‑specific currency formatting to Excel cells with Aspose.Cells for .NET without altering thread culture
// AI Prompts: Load an Excel workbook, set workbook.Settings.CultureInfo to "de-DE", iterate through column B, and assign the built‑in German currency number format (Number = 164) to each numeric cell. | Using Aspose.Cells for .NET, format all numeric cells in a worksheet as Euro currency based on the German locale while keeping the current thread culture unchanged. | Retrieve each cell's style, set Style.Number to 164 for the German € format, and save the workbook without modifying thread settings.
// Common Searches: aspnet aspocells format column as german euro currency without changing thread culture | set workbook culture to de-de and apply currency number format index 164 in Aspose.Cells | apply german locale number format to excel cells using Aspose.Cells C# | how to use built‑in number format 164 for € in Aspose.Cells workbook | preserve thread culture while formatting cells with german currency in Aspose.Cells
// Tags: Aspose.Cells workbook culture configuration | German Euro number style | locale‑specific cell formatting .NET | thread‑independent culture setting | format column B as currency

// Load the workbook (the file path can be adjusted as needed)
Aspose.Cells.LoadOptions loadOptions = new Aspose.Cells.LoadOptions();
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("input.xlsx", loadOptions);

// Apply German culture to the workbook (does not affect the current thread)
workbook.Settings.CultureInfo = new System.Globalization.CultureInfo("de-DE");

// Example: apply currency formatting to all numeric cells in column B of the first worksheet
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];
int firstRow = sheet.Cells.MinRow;
int lastRow = sheet.Cells.MaxRow;
int columnIndex = 1; // Column B (zero‑based index)

for (int row = firstRow; row <= lastRow; row++)
{
    Aspose.Cells.Cell cell = sheet.Cells[row, columnIndex];
    if (cell.Type == Aspose.Cells.CellValueType.IsNumeric)
    {
        // Use the built‑in German currency format (index 164 corresponds to "€ #,##0.00")
        Aspose.Cells.Style style = cell.GetStyle();
        style.Number = 164;
        cell.SetStyle(style);
    }
}

// Save the modified workbook
workbook.Save("output.xlsx");
