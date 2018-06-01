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
    /// Gets a value indicating if the file is a valid type.
    /// </summary>
    /// <param name="fileStream">A file stream.</param>
    /// <returns>If the file type is valid, true. Otherwise, false.</returns>
    public static bool IsValidFileType(Stream fileStream)
    {
      if (fileStream == null)
      {
        throw new ArgumentNullException("fileStream");
      }

      return IsFileHeaderValid(fileStream);
    }

    /// <summary>
    /// Gets a value indicating if the file is a valid type.
    /// </summary>
    /// <param name="fileBytes">The byte representation of a file.</param>
    /// <returns>If the file type is valid, true. Otherwise, false.</returns>
    public static bool IsValidFileType(byte[] fileBytes)
    {
      if (fileBytes == null)
      {
        throw new ArgumentNullException("fileBytes");
      }

      using (var stream = new MemoryStream(fileBytes, false))
      {
        return IsFileHeaderValid(stream);
      }
    }

    /// <summary>
    /// Gets a value indicating if the file is a valid type.
    /// </summary>
    /// <param name="filePath">A file path.</param>
    /// <returns>If the file type is valid, true. Otherwise, false.</returns>
    public static bool IsValidFileType(string filePath)
    {
      if (filePath == null)
      {
        throw new ArgumentNullException("filePath");
      }

      if (!File.Exists(filePath))
      {
        throw new FileNotFoundException();
      }

      using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
      {
        return IsFileHeaderValid(stream);
      }
    }

    /// <summary>
    /// Extracts text from a file.
    /// </summary>
    /// <param name="fileStream">A file stream.</param>
    /// <returns>The text content of a file.</returns>
    public static string Extract(Stream fileStream)
    {
      if (fileStream == null)
      {
        throw new ArgumentNullException("fileStream");
      }

      if (!IsFileHeaderValid(fileStream))
      {
        throw new NotSupportedException();
      }

      string content = ExtractFileContent(fileStream); ;

      // Don't bother extracting the text if the contents is blank.
      if (string.IsNullOrWhiteSpace(content))
      {
        return content;
      }

      return ExtractText(content);
    }

    /// <summary>
    /// Extracts text from a file.
    /// </summary>
    /// <param name="fileBytes">The byte representation of a file.</param>
    /// <returns>The text content of a file.</returns>
    public static string Extract(byte[] fileBytes)
    {
      if (fileBytes == null)
      {
        throw new ArgumentNullException("fileBytes");
      }

      using (var stream = new MemoryStream(fileBytes, false))
      {
        return Extract(stream);
      }
    }

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

      using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
      {
        return Extract(stream);
      }
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
          bool preserveSpace =
            match.Groups[2].Value.Equals("preserve", StringComparison.OrdinalIgnoreCase);

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

    private static string ExtractFileContent(Stream fileStream)
    {
      if (fileStream == null)
      {
        throw new ArgumentNullException("fileStream");
      }

      string content = string.Empty;

      using (var archive = new ZipArchive(fileStream, ZipArchiveMode.Read, true))
      {
        var document = archive.GetEntry(@"word/document.xml");

        if (document == null)
        {
          throw new NotSupportedException();
        }

        using (var documentStream = document.Open())
        {
          using (var reader = new StreamReader(documentStream))
          {
            content = reader.ReadToEnd();
          }
        }
      }

      return content;
    }

    private static bool IsFileHeaderValid(Stream fileStream)
    {
      if (fileStream == null)
      {
        throw new ArgumentNullException("fileStream");
      }

      var position = fileStream.Position;

      // This is the identifier for .docx.
      var identifier = new byte[] { 80, 75, 3, 4 };

      var header = new byte[identifier.Length];

      if (fileStream.Length >= header.Length)
      {
        if (fileStream.Read(header, 0, header.Length) != header.Length)
        {
          throw new EndOfStreamException();
        }
      }

      // Set the stream back to the position we found it.
      fileStream.Seek(position, SeekOrigin.Begin);

      if (!header.SequenceEqual(identifier))
      {
        return false;
      }

      return true;
    }
  }
}