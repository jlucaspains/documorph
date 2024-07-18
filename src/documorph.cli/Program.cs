var rootCommand = new RootCommand();
var markdownCommand = new Command("md", "Converts a .docx file to markdown.");

var sourceFileOption = new Option<FileInfo>(
            name: "--in",
            description: "The .docx file or directory to convert.")
{ IsRequired = true };

var targetFileOption = new Option<FileInfo>(
            name: "--out",
            description: "The output file or directory name.")
{ IsRequired = true };

rootCommand.AddCommand(markdownCommand);

markdownCommand.AddOption(sourceFileOption);
markdownCommand.AddOption(targetFileOption);

markdownCommand.SetHandler(Convert,
            sourceFileOption, targetFileOption);

return await rootCommand.InvokeAsync(args);

static async Task<int> Convert(FileInfo source, FileInfo target)
{
    Utilities.WriteInformation("Converting {0} to {1}", source.FullName, target.FullName);

    var isSourceDirectory = (source.Attributes & FileAttributes.Directory) == FileAttributes.Directory;

    if (!source.Exists && !isSourceDirectory)
    {
        Utilities.WriteError("The source file {0} does not exist.", source.FullName);
        return 1;
    }

    var outputDirectory = Path.GetDirectoryName(target.FullName)
            ?? "";
    
    if (isSourceDirectory && outputDirectory != target.FullName.TrimEnd('/').TrimEnd('\\'))
    {
        Utilities.WriteError("When source is a directory, target must also be a directory.");
        return 1;
    }

    var filesToProcess = isSourceDirectory
        ? Directory.GetFiles(source.FullName, "*.docx", SearchOption.TopDirectoryOnly).Select(item => new FileInfo(item))
        : [source];

    if (!Directory.Exists(outputDirectory))
    {
        Utilities.WriteInformation("Creating output directory {0}.", outputDirectory);
        Directory.CreateDirectory(outputDirectory);
    }

    foreach (var fileToProcess in filesToProcess)
    {
        var targetFile = outputDirectory == target.FullName || isSourceDirectory
            ? new FileInfo(Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(fileToProcess.FullName) + ".md"))
            : target;

        await ConvertFile(fileToProcess, targetFile, outputDirectory);
    }

    Utilities.WriteSuccess("Conversion completed successfully.");

    return 0;
}

static async Task ConvertFile(FileInfo fileToProcess, FileInfo targetFile, string mediaDirectory)
{
    Utilities.WriteInformation("Converting {0}...", fileToProcess.FullName);

    var processor = new DocxToMarkdownProcessor(fileToProcess.FullName);
    var (markdown, media) = processor.Process();

    using var markdownFileStream = targetFile.CreateText();
    await markdownFileStream.WriteAsync(markdown);

    if (!media.Any())
    {
        Utilities.WriteInformation("There were no media files to save.");
    }
    else
    {
        Utilities.WriteInformation("Saving media files to {0}. {1} media files will be saved.", mediaDirectory, media.Count());
    }

    foreach (var item in media)
    {
        var mediaPath = Path.Combine(mediaDirectory, item.UniqueFileName);
        var directory = Path.GetDirectoryName(mediaPath);

        await File.WriteAllBytesAsync(mediaPath, item.Content);
    }
}