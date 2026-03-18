using System;
using System.IO;
using System.Text;
using Aspose.Cells;

public class CustomStreamProvider : IStreamProvider
{
    private readonly string _htmlContent;
    private readonly string _outputDirectory;

    public CustomStreamProvider(string htmlContent, string outputDirectory)
    {
        _htmlContent = htmlContent;
        _outputDirectory = outputDirectory;
    }

    public void InitStream(StreamProviderOptions options)
    {
        if (options.DefaultPath != null && options.DefaultPath.EndsWith(".html", StringComparison.OrdinalIgnoreCase))
        {
            byte[] bytes = Encoding.UTF8.GetBytes(_htmlContent);
            options.Stream = new MemoryStream(bytes);
        }
        else
        {
            string filePath = Path.Combine(_outputDirectory, options.DefaultPath ?? "output.xlsx");
            Directory.CreateDirectory(Path.GetDirectoryName(filePath));
            options.Stream = File.Create(filePath);
        }
    }

    public void CloseStream(StreamProviderOptions options)
    {
        options.Stream?.Close();
        options.Stream = null;
    }
}

public class HtmlToXlsxViaStreamProvider
{
    public static void Run()
    {
        string html = @"
            <html>
                <body>
                    <table>
                        <tr><th>Id</th><th>Name</th></tr>
                        <tr><td>1</td><td>Alice</td></tr>
                        <tr><td>2</td><td>Bob</td></tr>
                    </table>
                </body>
            </html>";

        string outputDir = Path.Combine(Path.GetTempPath(), "AsposeStreamDemo");
        Directory.CreateDirectory(outputDir);

        var provider = new CustomStreamProvider(html, outputDir);

        // Load HTML from a memory stream using the provider
        using (var htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(html)))
        {
            var loadOptions = new HtmlLoadOptions
            {
                StreamProvider = provider
            };

            var workbook = new Workbook(htmlStream, loadOptions);

            // Prepare save options and initialise the output stream via the provider
            var saveOptions = new StreamProviderOptions(ResourceLoadingType.Default, "Result.xlsx");
            provider.InitStream(saveOptions);

            workbook.Save(saveOptions.Stream, SaveFormat.Xlsx);

            provider.CloseStream(saveOptions);
        }

        Console.WriteLine($"Workbook saved to: {Path.Combine(outputDir, "Result.xlsx")}");
    }
}

class Program
{
    static void Main()
    {
        HtmlToXlsxViaStreamProvider.Run();
    }
}