using System.CommandLine;
using documorph;

var rootCommand = new RootCommand();

var sourceFileOption = new Option<FileInfo>(
            name: "--in",
            description: "The file to convert. Can be .docx or .md.")
{ IsRequired = true };

var targetFileOption = new Option<FileInfo>(
            name: "--out",
            description: "The output file name. Can be .docx or .md.")
{ IsRequired = true };

rootCommand.AddOption(sourceFileOption);
rootCommand.AddOption(targetFileOption);

rootCommand.SetHandler(Convert,
            sourceFileOption, targetFileOption);

return await rootCommand.InvokeAsync(args);

static async Task<int> Convert(FileInfo source, FileInfo target)
{
    var processor = new DocxToMarkdownProcessor(source.FullName);
    var (markdown, media) = processor.Process();
    using var markdownFileStream = target.CreateText();
    await markdownFileStream.WriteAsync(markdown);

    var mediaDirectory = Path.GetDirectoryName(source.FullName)
        ?? "";

    foreach (var item in media)
    {
        var mediaPath = Path.Combine(mediaDirectory, item.FileName);
        var directory = Path.GetDirectoryName(mediaPath);
        
        if (directory != null)
            Directory.CreateDirectory(directory);

        await File.WriteAllBytesAsync(mediaPath, item.Content);
    }

    return 0;
}