namespace documorph.tests;

public class DocxToMarkdownTests
{
    [Fact]
    public void ValidDocxReturnsValidMarkdown()
    {
        var converter = new DocxToMarkdownProcessor("test_data/docconverter.docx");
        var (result, media) = converter.Process();

        Assert.NotNull(result);
        Assert.NotNull(media);
        Assert.NotNull(media);
        Assert.Contains("# Heading 1", result);
        Assert.Contains("## Heading 2", result);
        Assert.Contains("### Heading 3", result);
    }

    [Fact]
    public void NonExistingFileThrowsError()
    {
        var converter = new DocxToMarkdownProcessor("test_data/invalid.docx");
        Assert.Throws<FileNotFoundException>(() => converter.Process());
    }

    [Fact]
    public void BadFileThrowsError()
    {
        var converter = new DocxToMarkdownProcessor("test_data/badfile.docx");
        Assert.Throws<InvalidDataException>(() => converter.Process());
    }
}