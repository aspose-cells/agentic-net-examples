using System;
using Aspose.Cells;

public class MarkerHelper
{
    public string? GetValue(string name)
    {
        // Load the workbook (ensure the file path is correct)
        Workbook workbook = new Workbook("template.xlsx");

        // Initialize the designer with the loaded workbook
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Retrieve all smart markers in the workbook
        string[] smartMarkers = designer.GetSmartMarkers();

        // Find the first marker containing the specified name (case‑insensitive)
        foreach (string marker in smartMarkers)
        {
            if (!string.IsNullOrEmpty(marker) &&
                marker.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return marker;
            }
        }

        // No matching marker found
        return null;
    }
}

class Program
{
    static void Main()
    {
        var helper = new MarkerHelper();
        string? result = helper.GetValue("COLORS");
        Console.WriteLine(result ?? "Marker not found");
    }
}