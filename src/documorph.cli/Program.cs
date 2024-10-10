var rootCommand = new RootCommand();
var markdownCommand = new Command("md", "Converts a .docx file to markdown.");

// Add option for source file or directory
var sourceFileOption = new Option<FileInfo>(
            name: "--in",
            description: "The .docx file or directory to convert.")
{ IsRequired = true };

// Add option for target file or directory
var targetFileOption = new Option<FileInfo>(
            name: "--out",
            description: "The output file or directory name.")
{ IsRequired = true };

// Add option for unique media name option
var mediaLocationOption = new Option<DirectoryInfo>(
            name: "--media-location",
            description: "Media output directory.")
{ IsRequired = false };

rootCommand.AddCommand(markdownCommand);

markdownCommand.AddOption(sourceFileOption);
markdownCommand.AddOption(targetFileOption);
markdownCommand.AddOption(mediaLocationOption);

markdownCommand.SetHandler(Convert,
            sourceFileOption, targetFileOption, mediaLocationOption);

return await rootCommand.InvokeAsync(args);

static async Task<int> Convert(FileInfo source, FileInfo target,
    DirectoryInfo? mediaLocation)
{
    Utilities.WriteInformation("Converting {0} to {1}", source.FullName, target.FullName);

    var isSourceDirectory = (source.Attributes & FileAttributes.Directory) == FileAttributes.Directory;

    if (!source.Exists && !isSourceDirectory)
    {
        Utilities.WriteError("The source file {0} does not exist.", source.FullName);
        return 1;
    }

    var fileOutputDirectory = Path.GetDirectoryName(target.FullName)
            ?? "";

    if (isSourceDirectory && fileOutputDirectory != target.FullName.TrimEnd('/').TrimEnd('\\'))
    {
        Utilities.WriteError("When source is a directory, target must also be a directory.");
        return 1;
    }

    var filesToProcess = isSourceDirectory
        ? Directory.GetFiles(source.FullName, "*.docx", SearchOption.TopDirectoryOnly).Select(item => new FileInfo(item))
        : [source];

    if (!Directory.Exists(fileOutputDirectory))
    {
        Utilities.WriteInformation("Creating output directory {0}.", fileOutputDirectory);
        Directory.CreateDirectory(fileOutputDirectory);
    }

    var mediaOutputDirectory = mediaLocation ?? new DirectoryInfo(fileOutputDirectory);

    if (!mediaOutputDirectory.Exists)
    {
        Utilities.WriteInformation("Creating media output directory {0}.", fileOutputDirectory);
        mediaOutputDirectory.Create();
    }

    foreach (var fileToProcess in filesToProcess)
    {
        var targetFile = fileOutputDirectory == target.FullName || isSourceDirectory
            ? new FileInfo(Path.Combine(fileOutputDirectory, Path.GetFileNameWithoutExtension(fileToProcess.FullName) + ".md"))
            : target;

        await ConvertFile(fileToProcess, targetFile, mediaOutputDirectory);
    }

    Utilities.WriteSuccess("Conversion completed successfully.");

    return 0;
}

static async Task ConvertFile(FileInfo fileToProcess, FileInfo targetFile,
    DirectoryInfo mediaDirectory)
{
    Utilities.WriteInformation("Converting {0}...", fileToProcess.FullName);

    var targetFileDirectory = targetFile.DirectoryName ?? mediaDirectory.FullName;
    var mediaOutputRelativePath = Path.GetRelativePath(targetFileDirectory, mediaDirectory.FullName);

    var processor = new DocxToMarkdownProcessor(fileToProcess.FullName, mediaOutputRelativePath);
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
        var mediaPath = Path.Combine(targetFileDirectory, item.UniqueFileName);

        await File.WriteAllBytesAsync(mediaPath, item.Content);
    }
}
