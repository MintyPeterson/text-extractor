using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;

namespace MintyPeterson.TextExtractor
{
  /// <summary>
  /// Provides methods for extracting text from a file.
  /// </summary>
  public static class TextExtractor
  {
    /// <summary>
    /// Extracts text from a file.
    /// </summary>
    /// <param name="filePath">A file path.</param>
    /// <returns>The text content of a file.</returns>
    public static string Extract(string filePath)
    {
      if (filePath == null)
      {
        throw new ArgumentNullException("filePath");
      }

      if (!File.Exists(filePath))
      {
        throw new FileNotFoundException();
      }

      if(!IsFileHeaderValid(filePath))
      {
        throw new NotSupportedException();
      }

      var content = ExtractFileContent(filePath);

      // Don't bother extracting the text if the contents is blank.
      if (string.IsNullOrWhiteSpace(content))
      {
        return content;
      }

      return ExtractText(content);
    }

    private static string ExtractText(string contents)
    {
      var text = string.Empty;

      var pattern =
        @"(?:(<w:p\s.*?>|<w:br/>)|<w:t(?:(?:\sxml:space=""(preserve)"")|\s.*?|)>(.*?)<\/w:t>)";

      foreach (Match match in Regex.Matches(contents, pattern, RegexOptions.IgnoreCase))
      {
        if (match.Groups[1].Success)
        {
         // Add new lines as a single space.
          text += " ";
        }
        else
        {
          bool preserveSpace = match.Groups[2].Value.Equals("preserve", StringComparison.OrdinalIgnoreCase);

          if (preserveSpace)
          {
            text += match.Groups[3].Value;
          }
          else
          {
            text += match.Groups[3].Value.Trim();
          }
        }
      }

      return text.Trim();
    }

    private static string ExtractFileContent(string filePath)
    {
      string content = string.Empty;

      using (var archive = ZipFile.OpenRead(filePath))
      {
        var document = archive.GetEntry(@"word/document.xml");

        if (document == null)
        {
          throw new NotSupportedException();
        }

        using (var stream = document.Open())
        {
          using (var reader = new StreamReader(stream))
          {
            content = reader.ReadToEnd();
          }
        }
      }

      return content;
    }

    private static bool IsFileHeaderValid(string filePath)
    {
      // This is the identifier for .docx.
      var identifier = new byte[] { 80, 75, 3, 4 };

      var header = new byte[identifier.Length];

      using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
      {
        if (stream.Length >= header.Length)
        {
          if (stream.Read(header, 0, header.Length) != header.Length)
          {
            throw new EndOfStreamException();
          }
        }
      }

      if (!header.SequenceEqual(identifier))
      {
        return false;
      }

      return true;
    }
  }
}