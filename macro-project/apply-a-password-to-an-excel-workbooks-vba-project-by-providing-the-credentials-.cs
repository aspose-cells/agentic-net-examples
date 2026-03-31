using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Vba;

public class VbaProtectionInfo
{
    public bool IsLockedForViewing { get; set; }
    public string Password { get; set; }
}

public class Program
{
    public static void Main()
    {
        // 1. Read JSON credentials (create a default file if it does not exist)
        string jsonPath = "vba_protect.json";
        if (!File.Exists(jsonPath))
        {
            var defaultInfo = new VbaProtectionInfo
            {
                IsLockedForViewing = true,
                Password = "mySecretPwd"
            };
            string defaultJson = JsonSerializer.Serialize(defaultInfo, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(jsonPath, defaultJson);
        }

        string jsonContent = File.ReadAllText(jsonPath);
        VbaProtectionInfo protectInfo = JsonSerializer.Deserialize<VbaProtectionInfo>(jsonContent);

        // 2. Load the macro‑enabled workbook
        string inputWorkbookPath = "input.xlsm";
        if (!File.Exists(inputWorkbookPath))
        {
            Console.WriteLine($"Input workbook not found: {inputWorkbookPath}");
            return;
        }

        Workbook workbook = new Workbook(inputWorkbookPath);

        // 3. Apply protection to the VBA project if it exists
        if (workbook.VbaProject != null)
        {
            workbook.VbaProject.Protect(protectInfo.IsLockedForViewing, protectInfo.Password);
        }
        else
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
        }

        // 4. Save the protected workbook
        string outputWorkbookPath = "output_protected.xlsm";
        workbook.Save(outputWorkbookPath, SaveFormat.Xlsm);
        Console.WriteLine($"Workbook saved to {outputWorkbookPath}");
    }
}