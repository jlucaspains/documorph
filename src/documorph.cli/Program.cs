var rootCommand = new RootCommand();

var sourceFileOption = new Option<FileInfo>(
            name: "--in",
            description: "The .docx file to convert.")
{ IsRequired = true };

var targetFileOption = new Option<FileInfo>(
            name: "--out",
            description: "The output .md file name.")
{ IsRequired = true };

rootCommand.AddOption(sourceFileOption);
rootCommand.AddOption(targetFileOption);

rootCommand.SetHandler(Convert,
            sourceFileOption, targetFileOption);

return await rootCommand.InvokeAsync(args);

static async Task<int> Convert(FileInfo source, FileInfo target)
{
    Utilities.WriteInformation("Converting {0} to {1}", source.FullName, target.FullName);

    if (!source.Exists)
    {
        Utilities.WriteError("The source file {0} does not exist.", source.FullName);
        return 1;
    }

    var processor = new DocxToMarkdownProcessor(source.FullName);
    var (markdown, media) = processor.Process();

    var outputDirectory = Path.GetDirectoryName(target.FullName)
            ?? "";

    if (string.IsNullOrEmpty(outputDirectory))
    {
        Utilities.WriteError("The output directory is invalid.");
        return 2;
    }

    if (!Directory.Exists(outputDirectory))
    {
        Utilities.WriteInformation("Creating output directory {0}.", outputDirectory);
        Directory.CreateDirectory(outputDirectory);
    }

    using var markdownFileStream = target.CreateText();
    await markdownFileStream.WriteAsync(markdown);

    if (!media.Any())
    {
        Utilities.WriteInformation("There were no media files to save.");
        return 0;
    }

    Utilities.WriteInformation("Saving media files to {0}. {1} media files will be saved.", outputDirectory, media.Count());

    foreach (var item in media)
    {
        var mediaPath = Path.Combine(outputDirectory, item.FileName);
        var directory = Path.GetDirectoryName(mediaPath);

        await File.WriteAllBytesAsync(mediaPath, item.Content);
    }

    Utilities.WriteSuccess("Conversion completed successfully.");

    return 0;
}