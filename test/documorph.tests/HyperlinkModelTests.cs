
namespace documorph.tests;

public class HyperlinkModelTests
{
    [Fact]
    public void AppendMarkdown_WhenHyperlinkIdIsNull_ShouldNotIncludeHyperlinkUrl()
    {
        // Arrange
        var hyperlink = new Hyperlink();
        var hyperlinkRelationships = new List<HyperlinkRelationship>();

        var hyperlinkModel = new HyperlinkModel(hyperlink, hyperlinkRelationships);
        var builder = new StringBuilder();

        // Act
        hyperlinkModel.AppendMarkdown(builder);

        // Assert
        var result = builder.ToString();
        Assert.DoesNotContain("http", result);
    }

    [Fact]
    public void AppendMarkdown_WhenHyperlinkIdIsNotNull_ShouldIncludeHyperlinkUrl()
    {
        // Arrange
        var hyperlink = new Hyperlink(@"
<w:hyperlink 
xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" 
xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
r:id=""rId5"" w:history=""1"">
    <w:r w:rsidR=""00524EF3"" w:rsidRPr=""00524EF3"">
        <w:rPr>
            <w:rStyle w:val=""Hyperlink""/>
        </w:rPr>
        <w:t>A link</w:t>
    </w:r>
</w:hyperlink>");

        using var wordDoc = WordprocessingDocument.Create("test.docx", WordprocessingDocumentType.Document);
        wordDoc.AddMainDocumentPart();
        var hyperlinkRel = wordDoc.MainDocumentPart!.AddHyperlinkRelationship(new Uri("https://example.com"), true, "rId5");

        var hyperlinkModel = new HyperlinkModel(hyperlink, [hyperlinkRel]);
        var builder = new StringBuilder();

        // Act
        hyperlinkModel.AppendMarkdown(builder);

        // Assert
        var result = builder.ToString();
        Assert.Equal("[A link](https://example.com/)", result);
    }
    
    [Fact]
    public void AppendMarkdown_WillEscapeValue_ShouldIncludeHyperlinkUrl()
    {
        // Arrange
        var hyperlink = new Hyperlink(@"
<w:hyperlink 
xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" 
xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
r:id=""rId5"" w:history=""1"">
    <w:r w:rsidR=""00524EF3"" w:rsidRPr=""00524EF3"">
        <w:rPr>
            <w:rStyle w:val=""Hyperlink""/>
        </w:rPr>
        <w:t>(A) [link]</w:t>
    </w:r>
</w:hyperlink>");

        using var wordDoc = WordprocessingDocument.Create("test.docx", WordprocessingDocumentType.Document);
        wordDoc.AddMainDocumentPart();
        var hyperlinkRel = wordDoc.MainDocumentPart!.AddHyperlinkRelationship(new Uri("https://example.com"), true, "rId5");

        var hyperlinkModel = new HyperlinkModel(hyperlink, [hyperlinkRel]);
        var builder = new StringBuilder();

        // Act
        hyperlinkModel.AppendMarkdown(builder);

        // Assert
        var result = builder.ToString();
        Assert.Equal("[\\(A\\) \\[link\\]](https://example.com/)", result);
    }
}