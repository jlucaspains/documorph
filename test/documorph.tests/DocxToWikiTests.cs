using System.Text.RegularExpressions;

namespace documorph.tests;

public partial class DocxToWikiTests
{
    private const string EXPECTED_SAMPLE = @"= Heading 1 =

Some first paragraph. 

'''Some bold'''.

''Some italics''.

[https://blog.lpains.net/ A link].

== Heading 2 ==

Another paragraph.

=== Heading 3 ===

* Bullet 1
** Child Bullet 1.1
*** Child Bullet 1.1.1
* Bullet 2
* Bullet 3

==== Heading 4 ====

# List 1
## Child List 1.1
### Child Bullet 1.1.1
# List 2
# List 3

{|
|Column 1
|Column 2
|Column 3
|Column 4
|Column 5
|-
|Cel 1 1
|Cell 2 1
|Cell 3 1
|Cell 4 1
|Cell 5 1
|-
|Cell 1 2
|Cell 2 2 
|Cell 3 2
|Cell 4 2
|Cell 5 2
|-
|Cell 1 3
|Cell 2 3
|Cell 3 3
|Cell 4 3
|Cell 4 3
|}
<blockquote>This is a quote</blockquote>

<blockquote>This is a quote2</blockquote>

<blockquote>This is a quote3</blockquote>

[[File:./image1.png|A screenshot of a computer  Description automatically generated]]




After a break

";

    [Fact]
    public void ValidDocxReturnsValidMarkdown()
    {
        var converter = new DocxToWikiProcessor("test_data/docconverter.docx", ".");
        var (result, media) = converter.Process();

        // code files are saved using windows format.
        // To ensure tests work on both environments, replace \r\n with environment new line 
        var expected = EXPECTED_SAMPLE
            .Replace("\r\n", Environment.NewLine);

        var deterministicResult = GuidRegex().Replace(result, "image1.png");

        Assert.NotNull(result);
        Assert.NotNull(media);
        Assert.NotNull(media);
        Assert.Equal(expected, deterministicResult);
    }

    [Fact]
    public void ValidDocxWithMediaRelativePathReturnsValidMarkdown()
    {
        var converter = new DocxToWikiProcessor("test_data/docconverter.docx", "../../");
        var (result, media) = converter.Process();

        // code files are saved using windows format.
        // To ensure tests work on both environments, replace \r\n with environment new line 
        var expected = EXPECTED_SAMPLE
            .Replace("\r\n", Environment.NewLine)
            .Replace("./image1.png", "../../image1.png");

        var deterministicResult = GuidRegex().Replace(result, "image1.png");

        Assert.NotNull(result);
        Assert.NotNull(media);
        Assert.NotNull(media);
        Assert.Equal(expected, deterministicResult);
    }

    [Fact]
    public void NonExistingFileThrowsError()
    {
        var converter = new DocxToWikiProcessor("test_data/invalid.docx", ".");
        Assert.Throws<FileNotFoundException>(() => converter.Process());
    }

    [Fact]
    public void BadFileThrowsError()
    {
        var converter = new DocxToWikiProcessor("test_data/badfile.docx", ".");
        Assert.Throws<InvalidDataException>(() => converter.Process());
    }
    
    [Fact]
    public void WillMoveSpacesOutOfBold()
    {
        var converter = new DocxToWikiProcessor("test_data/boldwithextraspace.docx", ".");
        var expected = $"'''This is a bold string.''' This is not.{Environment.NewLine}{Environment.NewLine}";
        var (result, _) = converter.Process();
        Assert.Equal(expected, result);
    }

    [Fact]
    public void WillCorrectlyHandleBoldAndItalics()
    {
        var converter = new DocxToWikiProcessor("test_data/boldanditalics.docx", ".");
        var expected = $"'''''Bold and Italics'''''{Environment.NewLine}{Environment.NewLine}'''<u>Bold and Underline</u>'''{Environment.NewLine}{Environment.NewLine}";
        var (result, _) = converter.Process();
        Assert.Equal(expected, result);
    }

    [GeneratedRegex("([0-9A-Fa-f]{8}[-]?[0-9A-Fa-f]{4}[-]?[0-9A-Fa-f]{4}[-]?[0-9A-Fa-f]{4}[-]?[0-9A-Fa-f]{12}).png")]
    private static partial Regex GuidRegex();
}