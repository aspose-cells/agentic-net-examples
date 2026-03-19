using System;
using Aspose.Cells;
using Aspose.Cells.Vba;
using System.Text.Json;

public class VbaProjectStatus
{
    public bool IsProtected { get; set; }
    public bool IsLockedForViewing { get; set; }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Path to the Excel file (macro-enabled). Use first argument or default.
        string filePath = args.Length > 0 ? args[0] : "sample.xlsm";

        // Load the workbook (lifecycle rule)
        Workbook workbook = new Workbook(filePath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Gather protection status
        VbaProjectStatus status = new VbaProjectStatus
        {
            IsProtected = vbaProject.IsProtected,
            IsLockedForViewing = vbaProject.IslockedForViewing
        };

        // Convert status to JSON
        string json = JsonSerializer.Serialize(status);

        // Output the JSON string
        Console.WriteLine(json);
    }
}