# documorph
`documorph` is a .NET package and command-line tool for converting between document file formats. The initial implementation supports only `.docx` to `.md` files, but other formats will be considered for the future.

## Package
```powershell
dotnet package install lpains.documorph --prerelease
```

### Getting Started
```csharp
// Create an instance of the DocxToMarkdownProcessor class. This class requires the .docx file path.
var processor = new DocxToMarkdownProcessor(source.FullName);

// Invoke the Process() method which returns the markdown content and media files.
var (markdown, media) = processor.Process();
```

## CLI
```powershell
dotnet tool install --global lpains.documorph.cli --prerelease
```

### Getting Started
Upon installation, access the tool by executing `documorph` in your terminal. For specific command details, refer to the sections below or utilize the CLI help via `documorph -h`.

```powershell
documorph --in <inputFile> --out <outputFile> [-?, -h, --help]
```

Basic usage example:
```powershell
documorph --in .\source.docx `
           --out .\target.md
```

Output file (target.md):
```markdown
# Heading 1

1. numbered lists are supported

## Heading 2

- bullet lists too

### Heading 3

> You can create quotes and tables

| Column 1 | Column 2 |
|----------|----------|
| value 1  | value 2  |

#### Heading 4
You can also add links like this: [Link](https://www.example.com)

And images like this:
![Image](image1.png)

And **bold** or *italic* or __underscore__ or ~~striked~~ text.
```

### Parameters
##### `--in` (required)

The input `.docx` file. It should be a valid Open XML Word document.

#### `--out` (required)

The output file name. It should be a file with `.md` extension.
