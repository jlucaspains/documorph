namespace documorph.tests;

public class DocxToMarkdownTests
{
    private const string EXPECTED_SAMPLE = @"# Heading 1

Some first paragraph. 

**Some bold**.

*Some italics*.

[A link](https://blog.lpains.net/).

## Heading 2

Another paragraph.

### Heading 3

- Bullet 1
   - Child Bullet 1.1
      - Child Bullet 1.1.1
- Bullet 2
- Bullet 3

#### Heading 4

1. List 1
   1. Child List 1.1
      1. Child Bullet 1.1.1
1. List 2
1. List 3

|Column 1|Column 2|Column 3|Column 4|Column 5|
|---|---|---|---|---|
|Cel 1 1|Cell 2 1|Cell 3 1|Cell 4 1|Cell 5 1|
|Cell 1 2|Cell 2 2 |Cell 3 2|Cell 4 2|Cell 5 2|
|Cell 1 3|Cell 2 3|Cell 3 3|Cell 4 3|Cell 4 3|

> This is a quote

> This is a quote2

> This is a quote3

![A screenshot of a computer  Description automatically generated](image1.png)

";

    [Fact]
    public void ValidDocxReturnsValidMarkdown()
    {
        var converter = new DocxToMarkdownProcessor("test_data/docconverter.docx");
        var (result, media) = converter.Process();

        Assert.NotNull(result);
        Assert.NotNull(media);
        Assert.NotNull(media);
        Assert.Equal(EXPECTED_SAMPLE, result);
        Assert.True(File.Exists("test_data/image1.png"));
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