using System.IO.Packaging;

namespace documorph.tests;

public class ParagraphModelTests
{
    [Fact]
    public void AppendMarkdown_Should_Append_HeaderMarkdownPrefix_When_IsHeading()
    {
        // Arrange
        var paragraph = new Paragraph(@"
<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml""
w14:paraId=""726C5E95"" w14:textId=""7029D8CB"" w:rsidR=""00253B1F"" w:rsidRDefault=""00FD6CC0"" w:rsidP=""00FD6CC0"">
    <w:pPr>
        <w:pStyle w:val=""Heading1""/>
    </w:pPr>
    <w:r>
        <w:t>Heading 1</w:t>
    </w:r>
</w:p>");

        // Act
        var paragraphModel = ParagraphModel.FromParagraph(paragraph, null, [], []);

        // Assert
        Assert.True(paragraphModel.IsHeading);
        Assert.False(paragraphModel.IsList);
        Assert.False(paragraphModel.IsQuote);
        Assert.Equal(1, paragraphModel.HeadingLevel);
        Assert.Equal("Heading 1", paragraphModel.Children.OfType<RunModel>().First().Text);
    }

    [Fact]
    public void AppendMarkdown_Should_Append_ListMarkdownPrefix_When_IsList()
    {
        // Arrange
        var paragraph = new Paragraph(@"
<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml""
w14:paraId=""4263CB73"" w14:textId=""2C2F97A4"" w:rsidR=""00FD6CC0"" w:rsidRDefault=""00FD6CC0"" w:rsidP=""00FD6CC0"">
    <w:pPr>
        <w:pStyle w:val=""ListParagraph""/>
        <w:numPr>
            <w:ilvl w:val=""0""/>
            <w:numId w:val=""1""/>
        </w:numPr>
    </w:pPr>
    <w:r>
        <w:t>Bullet 1</w:t>
    </w:r>
</w:p>");

        using var wordPackage = Package.Open("test_data/docconverter.docx", FileMode.Open, FileAccess.Read, FileShare.Read);
        using var wordDocument = WordprocessingDocument.Open(wordPackage);

        // Act
        var paragraphModel = ParagraphModel.FromParagraph(paragraph, wordDocument.MainDocumentPart!.NumberingDefinitionsPart, [], []);

        // Assert
        Assert.False(paragraphModel.IsHeading);
        Assert.True(paragraphModel.IsList);
        Assert.Equal("bullet", paragraphModel.ListFormat);
        Assert.Equal(0, paragraphModel.ListItemLevel);
        Assert.Equal("Bullet 1", paragraphModel.Children.OfType<RunModel>().First().Text);
    }

    [Fact]
    public void AppendMarkdown_Should_Append_ListMarkdown_When_IsListAndIdented()
    {
        // Arrange
        var paragraph = new Paragraph(@"
<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml""
w14:paraId=""4263CB73"" w14:textId=""2C2F97A4"" w:rsidR=""00FD6CC0"" w:rsidRDefault=""00FD6CC0"" w:rsidP=""00FD6CC0"">
    <w:pPr>
        <w:pStyle w:val=""ListParagraph""/>
        <w:numPr>
            <w:ilvl w:val=""1""/>
            <w:numId w:val=""1""/>
        </w:numPr>
    </w:pPr>
    <w:r>
        <w:t>Bullet 1</w:t>
    </w:r>
</w:p>");

        using var wordPackage = Package.Open("test_data/docconverter.docx", FileMode.Open, FileAccess.Read, FileShare.Read);
        using var wordDocument = WordprocessingDocument.Open(wordPackage);

        // Act
        var paragraphModel = ParagraphModel.FromParagraph(paragraph, wordDocument.MainDocumentPart!.NumberingDefinitionsPart, [], []);

        // Assert
        Assert.False(paragraphModel.IsHeading);
        Assert.False(paragraphModel.IsQuote);
        Assert.True(paragraphModel.IsList);
        Assert.Equal("bullet", paragraphModel.ListFormat);
        Assert.Equal(1, paragraphModel.ListItemLevel);
        Assert.Equal("Bullet 1", paragraphModel.Children.OfType<RunModel>().First().Text);
    }
    
    [Fact]
    public void AppendMarkdown_Should_Append_ListMarkdown_When_IsListAndList()
    {
        // Arrange
        var paragraph = new Paragraph(@"
<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml""
w14:paraId=""4263CB73"" w14:textId=""2C2F97A4"" w:rsidR=""00FD6CC0"" w:rsidRDefault=""00FD6CC0"" w:rsidP=""00FD6CC0"">
    <w:pPr>
        <w:pStyle w:val=""ListParagraph""/>
        <w:numPr>
            <w:ilvl w:val=""0""/>
            <w:numId w:val=""2""/>
        </w:numPr>
    </w:pPr>
    <w:r>
        <w:t>List 1</w:t>
    </w:r>
</w:p>");

        using var wordPackage = Package.Open("test_data/docconverter.docx", FileMode.Open, FileAccess.Read, FileShare.Read);
        using var wordDocument = WordprocessingDocument.Open(wordPackage);

        // Act
        var paragraphModel = ParagraphModel.FromParagraph(paragraph, wordDocument.MainDocumentPart!.NumberingDefinitionsPart, [], []);

        // Assert
        Assert.False(paragraphModel.IsHeading);
        Assert.False(paragraphModel.IsQuote);
        Assert.True(paragraphModel.IsList);
        Assert.Equal("decimal", paragraphModel.ListFormat);
        Assert.Equal(0, paragraphModel.ListItemLevel);
        Assert.Equal("List 1", paragraphModel.Children.OfType<RunModel>().First().Text);
    }
    
    [Fact]
    public void AppendMarkdown_Should_Append_QuoteMarkdownPrefix_When_IsQuote()
    {
        // Arrange
        var paragraph = new Paragraph(@"
<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml""
w14:paraId=""10F62EBF"" w14:textId=""1DC29406"" w:rsidR=""00C276F4"" w:rsidRDefault=""00C276F4"" w:rsidP=""00C276F4"">
    <w:pPr>
        <w:pStyle w:val=""Quote""/>
    </w:pPr>
    <w:r>
        <w:t xml:space=""preserve"">This is a </w:t>
    </w:r>
    <w:proofErr w:type=""gramStart""/>
    <w:r>
        <w:t>quote</w:t>
    </w:r>
    <w:proofErr w:type=""gramEnd""/>
</w:p>");


        // Act
        var paragraphModel = ParagraphModel.FromParagraph(paragraph, null, [], []);

        // Assert
        Assert.False(paragraphModel.IsHeading);
        Assert.True(paragraphModel.IsQuote);
        Assert.False(paragraphModel.IsList);
        Assert.Equal("This is a ", paragraphModel.Children.OfType<RunModel>().First().Text);
    }

    [Fact]
    public void AppendMarkdown_Should_Append_HyperlinkMarkdown_When_IsHyperlink()
    {
        // Arrange
        var paragraph = new Paragraph(@"
<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" 
xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml""
w14:paraId=""056E7E2B"" w14:textId=""0A0BA3C9"" w:rsidR=""00FD6CC0"" w:rsidRDefault=""00000000"" w:rsidP=""00FD6CC0"">
    <w:hyperlink 
    r:id=""rId5"" w:history=""1"">
        <w:r w:rsidR=""00524EF3"" w:rsidRPr=""00524EF3"">
            <w:rPr>
                <w:rStyle w:val=""Hyperlink""/>
            </w:rPr>
            <w:t>A link</w:t>
        </w:r>
    </w:hyperlink>
</w:p>");

        using var wordDoc = WordprocessingDocument.Create("AppendMarkdown_Should_Append_HyperlinkMarkdown_When_IsHyperlink.docx", WordprocessingDocumentType.Document);
        wordDoc.AddMainDocumentPart();
        var hyperlinkRel = wordDoc.MainDocumentPart!.AddHyperlinkRelationship(new Uri("https://example.com"), true, "rId5");

        // Act
        var paragraphModel = ParagraphModel.FromParagraph(paragraph, null, [hyperlinkRel], []);

        // Assert
        var hyperlinkModel = paragraphModel.Children.OfType<HyperlinkModel>().First();
        Assert.False(paragraphModel.IsHeading);
        Assert.False(paragraphModel.IsQuote);
        Assert.False(paragraphModel.IsList);
        Assert.Equal("https://example.com/", hyperlinkModel.Uri);
        Assert.Equal("A link", hyperlinkModel.Runs.First().Text);
    }
}
