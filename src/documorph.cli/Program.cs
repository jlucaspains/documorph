using DocumentFormat.OpenXml.Bibliography;

var rootCommand = new RootCommand();
var markdownCommand = new Command("md", "Converts a .docx file to markdown.");
var asciidocCommand = new Command("asciidoc", "Converts a .docx file to asciidoc.");
var wikiCommand = new Command("wiki", "Converts a .docx file to wikimedia format.");

// Add option for source file or directory
var sourceFileOption = new Option<FileInfo>(
            name: "--in")
{  Required = true, Description = "The .docx file or directory to convert." };

// Add option for target file or directory
var targetFileOption = new Option<FileInfo>(
            name: "--out")
{ Required = true, Description = "The output file or directory name." };

// Add option for unique media name option
var mediaLocationOption = new Option<DirectoryInfo>(
            name: "--media-location")
{ Required = false, Description = "Media output directory." };

rootCommand.Add(markdownCommand);
rootCommand.Add(asciidocCommand);
rootCommand.Add(wikiCommand);

markdownCommand.Options.Add(sourceFileOption);
markdownCommand.Options.Add(targetFileOption);
markdownCommand.Options.Add(mediaLocationOption);

asciidocCommand.Options.Add(sourceFileOption);
asciidocCommand.Options.Add(targetFileOption);
asciidocCommand.Options.Add(mediaLocationOption);

wikiCommand.Options.Add(sourceFileOption);
wikiCommand.Options.Add(targetFileOption);
wikiCommand.Options.Add(mediaLocationOption);

markdownCommand.SetAction(async parseResult => 
{
    var sourceFile = parseResult.GetValue(sourceFileOption);
    var targetFile = parseResult.GetValue(targetFileOption);
    var mediaLocation = parseResult.GetValue(mediaLocationOption);

    return await ConvertMarkdown(sourceFile, targetFile, mediaLocation);
});

asciidocCommand.SetAction(async parseResult => 
{
    var sourceFile = parseResult.GetValue(sourceFileOption);
    var targetFile = parseResult.GetValue(targetFileOption);
    var mediaLocation = parseResult.GetValue(mediaLocationOption);

    return await ConvertAsciiDoc(sourceFile, targetFile, mediaLocation);
});

wikiCommand.SetAction(async parseResult => 
{
    var sourceFile = parseResult.GetValue(sourceFileOption);
    var targetFile = parseResult.GetValue(targetFileOption);
    var mediaLocation = parseResult.GetValue(mediaLocationOption);

    return await ConvertWiki(sourceFile, targetFile, mediaLocation);
});

ParseResult parseResult = rootCommand.Parse(args);

return await parseResult.InvokeAsync();

static async Task<int> ConvertMarkdown(FileInfo? source, FileInfo? target,
    DirectoryInfo? mediaLocation)
{
    return await Convert(TargetType.Markdown, source, target, mediaLocation);
}

static async Task<int> ConvertAsciiDoc(FileInfo? source, FileInfo? target,
    DirectoryInfo? mediaLocation)
{
    return await Convert(TargetType.AsciiDoc, source, target, mediaLocation);
}

static async Task<int> ConvertWiki(FileInfo? source, FileInfo? target,
    DirectoryInfo? mediaLocation)
{
    return await Convert(TargetType.Wiki, source, target, mediaLocation);
}

static async Task<int> Convert(TargetType targetType, FileInfo? source, FileInfo? target,
    DirectoryInfo? mediaLocation)
{
    if (source == null)
    {
        Utilities.WriteError("The --source option is required.");
        return 1;
    }

    if (target == null)
    {
        Utilities.WriteError("The --target option is required.");
        return 1;
    }

    Utilities.WriteInformation("Converting {0} to {1}", source.FullName, target.FullName);

    var isSourceDirectory = (source.Attributes & FileAttributes.Directory) == FileAttributes.Directory;

    if (!source.Exists && !isSourceDirectory)
    {
        Utilities.WriteError("The source file {0} does not exist.", source.FullName);
        return 1;
    }

    var fileOutputDirectory = !isSourceDirectory ? Path.GetDirectoryName(target.FullName) : target.FullName;

    if (isSourceDirectory && File.Exists(fileOutputDirectory))
    {
        Utilities.WriteError("When source is a directory, target must also be a directory.");
        return 1;
    }

    var filesToProcess = isSourceDirectory
        ? Directory.GetFiles(source.FullName, "*.docx", SearchOption.TopDirectoryOnly).Select(item => new FileInfo(item))
        : [source];

    if (!Directory.Exists(fileOutputDirectory))
    {
        fileOutputDirectory ??= string.Empty;
        Utilities.WriteInformation("Creating output directory {0}.", fileOutputDirectory);
        Directory.CreateDirectory(fileOutputDirectory);
    }

    var mediaOutputDirectory = mediaLocation ?? new DirectoryInfo(fileOutputDirectory);

    if (!mediaOutputDirectory.Exists)
    {
        Utilities.WriteInformation("Creating media output directory {0}.", fileOutputDirectory);
        mediaOutputDirectory.Create();
    }

    var defaultExtension = targetType switch
    {
        TargetType.Markdown => ".md",
        TargetType.AsciiDoc => ".asciidoc",
        TargetType.Wiki => ".wiki",
        _ => throw new NotImplementedException(),
    };

    foreach (var fileToProcess in filesToProcess)
    {
        var targetFile = isSourceDirectory
            ? new FileInfo(Path.Combine(fileOutputDirectory, Path.GetFileNameWithoutExtension(fileToProcess.FullName) + defaultExtension))
            : target;

        await ConvertFile(targetType, fileToProcess, targetFile, mediaOutputDirectory);
    }

    Utilities.WriteSuccess("Conversion completed successfully.");

    return 0;
}

static async Task ConvertFile(TargetType targetType, FileInfo fileToProcess, FileInfo targetFile,
    DirectoryInfo mediaDirectory)
{
    Utilities.WriteInformation("Converting {0}...", fileToProcess.FullName);

    var targetFileDirectory = targetFile.DirectoryName ?? mediaDirectory.FullName;
    var mediaOutputRelativePath = Path.GetRelativePath(targetFileDirectory, mediaDirectory.FullName);

    IDocxToTargetConverter processor = targetType switch
    {
        TargetType.Markdown => new DocxToMarkdownProcessor(fileToProcess.FullName, mediaOutputRelativePath),
        TargetType.AsciiDoc => new DocxToAsciiDocProcessor(fileToProcess.FullName, mediaOutputRelativePath),
        TargetType.Wiki => new DocxToWikiProcessor(fileToProcess.FullName, mediaOutputRelativePath),
        _ => throw new InvalidOperationException("Invalid target type.")
    };
    // var processor = new DocxToMarkdownProcessor(fileToProcess.FullName, mediaOutputRelativePath);
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