using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Please provide the path to the workbook as a command‑line argument.");
            return;
        }

        string filePath = args[0];
        CheckVbaProtection.Run(filePath);
    }
}

static class CheckVbaProtection
{
    public static void Run(string filePath)
    {
        // Load the workbook
        Workbook workbook = new Workbook(filePath);

        // Get the VBA project (may be null if none exists)
        VbaProject? vbaProject = workbook.VbaProject;

        // Determine if a VBA project is present and whether it is protected
        bool hasVbaProject = vbaProject != null;
        bool isProtected = hasVbaProject && vbaProject.IsProtected;

        // Output the results
        Console.WriteLine($"Workbook: {filePath}");
        Console.WriteLine($"VBA project exists: {hasVbaProject}");
        Console.WriteLine($"VBA project protected: {isProtected}");
    }
}