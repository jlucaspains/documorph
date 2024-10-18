namespace documorph.tests;

public class ProgramMarkdownTests
{
    [Fact]
    public void MissingInReturn1()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "md", "--out", "test.md" }]);

        Assert.Equal(1, result);
    }

    [Fact]
    public void MissingOutReturn1()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "md", "--in", "test.docx" }]);

        Assert.Equal(1, result);
    }

    [Fact]
    public void WillConvertSimpleDocx()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "md", "--in", "./test_data/simpletest.docx", "--out", "./test_data/WillConvertSimpleDocx.md" }]);

        Assert.Equal(0, result);
        Assert.True(File.Exists("./test_data/WillConvertSimpleDocx.md"));
    }
    
    [Fact]
    public void ValidParameters()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "md", "--in", "./test_data/docconverter.docx", "--out", "./test_data/Test_ValidParameters.md" }]);

        Assert.Equal(0, result);
        Assert.True(File.Exists("./test_data/Test_ValidParameters.md"));
    }
    
    [Fact]
    public void WillCreateOutputDirectoryIfNeeded()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "md", "--in", "./test_data/docconverter.docx", "--out", "./test_data2/Test_ValidParameters.md" }]);

        Assert.Equal(0, result);
        Assert.True(File.Exists("./test_data2/Test_ValidParameters.md"));
    }
    
    [Fact]
    public void InvalidInFileParameters()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "md", "--in", "test_data/docconverterbad.docx", "--out", "test_data/Test_ValidParameters.md" }]);

        Assert.Equal(1, result);
    }
    
    [Fact]
    public void InvalidOutFileParameters()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "md", "--in", "test_data/docconverter.docx", "--out", "" }]);

        Assert.Equal(1, result);
    }
    
    [Fact]
    public void WillFailIfInputIsDirectoryAndOutputIsNotDirectory()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "md", "--in", "test_data/", "--out", "test_data/docconverter.docx" }]);

        Assert.Equal(1, result);
    }

    [Fact]
    public void WillConvertAnEntireFolder()
    {
        Directory.CreateDirectory("WillConvertAnEntireFolderMD/");
        File.Copy("test_data/docconverter.docx", "WillConvertAnEntireFolderMD/docconverter1.docx", true);
        File.Copy("test_data/docconverter.docx", "WillConvertAnEntireFolderMD/docconverter2.docx", true);
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "md", "--in", "WillConvertAnEntireFolderMD/", "--out", "WillConvertAnEntireFolderMD/" }]);

        Assert.Equal(0, result);
        Assert.Equal(Directory.GetFiles("WillConvertAnEntireFolderMD/", "*.docx").Length, 
            Directory.GetFiles("WillConvertAnEntireFolderMD/", "*.md").Length);
    }
}