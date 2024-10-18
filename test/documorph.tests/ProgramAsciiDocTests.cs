namespace documorph.tests;

public class ProgramAsciiDocTests
{
    [Fact]
    public void MissingInReturn1()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "asciidoc", "--out", "test.md" }]);

        Assert.Equal(1, result);
    }

    [Fact]
    public void MissingOutReturn1()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "asciidoc", "--in", "test.docx" }]);

        Assert.Equal(1, result);
    }

    [Fact]
    public void WillConvertSimpleDocx()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "asciidoc", "--in", "./test_data/simpletest.docx", "--out", "./test_data/WillConvertSimpleDocx.asciidoc" }]);

        Assert.Equal(0, result);
        Assert.True(File.Exists("./test_data/WillConvertSimpleDocx.asciidoc"));
    }
    
    [Fact]
    public void ValidParameters()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "asciidoc", "--in", "./test_data/docconverter.docx", "--out", "./test_data/Test_ValidParameters.asciidoc" }]);

        Assert.Equal(0, result);
        Assert.True(File.Exists("./test_data/Test_ValidParameters.asciidoc"));
    }
    
    [Fact]
    public void WillCreateOutputDirectoryIfNeeded()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "asciidoc", "--in", "./test_data/docconverter.docx", "--out", "./test_data2/Test_ValidParameters.asciidoc" }]);

        Assert.Equal(0, result);
        Assert.True(File.Exists("./test_data2/Test_ValidParameters.asciidoc"));
    }
    
    [Fact]
    public void InvalidInFileParameters()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "asciidoc", "--in", "test_data/docconverterbad.docx", "--out", "test_data/Test_ValidParameters.asciidoc" }]);

        Assert.Equal(1, result);
    }
    
    [Fact]
    public void InvalidOutFileParameters()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "asciidoc", "--in", "test_data/docconverter.docx", "--out", "" }]);

        Assert.Equal(1, result);
    }
    
    [Fact]
    public void WillFailIfInputIsDirectoryAndOutputIsNotDirectory()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "asciidoc", "--in", "test_data/", "--out", "test_data/docconverter.docx" }]);

        Assert.Equal(1, result);
    }

    [Fact]
    public void WillConvertAnEntireFolder()
    {
        Directory.CreateDirectory("WillConvertAnEntireFolderAsciiDoc/");
        File.Copy("test_data/docconverter.docx", "WillConvertAnEntireFolderAsciiDoc/docconverter1.docx", true);
        File.Copy("test_data/docconverter.docx", "WillConvertAnEntireFolderAsciiDoc/docconverter2.docx", true);
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "asciidoc", "--in", "WillConvertAnEntireFolderAsciiDoc/", "--out", "WillConvertAnEntireFolderAsciiDoc/" }]);

        Assert.Equal(0, result);
        Assert.Equal(Directory.GetFiles("WillConvertAnEntireFolderAsciiDoc/", "*.docx").Length, 
            Directory.GetFiles("WillConvertAnEntireFolderAsciiDoc/", "*.asciidoc").Length);
    }
}