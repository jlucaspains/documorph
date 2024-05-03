namespace documorph.tests;

public class ProgramTests
{
    [Fact]
    public void NoParametersReturn1()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [Array.Empty<string>()]);

        Assert.Equal(1, result);
    }

    [Fact]
    public void MissingInReturn1()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "--out", "test.md" }]);

        Assert.Equal(1, result);
    }

    [Fact]
    public void MissingOutReturn1()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "--in", "test.docx" }]);

        Assert.Equal(1, result);
    }

    [Fact]
    public void ValidParameters()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "--in", "test_data/docconverter.docx", "--out", "test_data/Test_ValidParameters.md" }]);

        Assert.Equal(0, result);
        Assert.True(File.Exists("test_data/Test_ValidParameters.md"));
        Assert.True(File.Exists("test_data/image1.png"));
    }
}