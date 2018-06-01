using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace MintyPeterson.TextExtractor.Tests
{
  [TestClass]
  public class TextExtractorUnitTest
  {
    [TestMethod]
    public void TestExtractMethodWithUndefinedFile()
    {
      Assert.ThrowsException<ArgumentNullException>(
        () => TextExtractor.Extract((string)null)
      );
    }

    [TestMethod]
    public void TestExtractMethodWithUndefinedStream()
    {
      Assert.ThrowsException<ArgumentNullException>(
        () => TextExtractor.Extract((Stream)null)
      );
    }

    [TestMethod]
    public void TestExtractMethodWithUndefinedBytes()
    {
      Assert.ThrowsException<ArgumentNullException>(
        () => TextExtractor.Extract((byte[])null)
      );
    }

    [TestMethod]
    public void TestExtractMethodWithMissingStream()
    {
      Assert.ThrowsException<NotSupportedException>(
        () => TextExtractor.Extract(new MemoryStream())
      );
    }

    [TestMethod]
    public void TestExtractMethodWithMissingBytes()
    {
      Assert.ThrowsException<NotSupportedException>(
        () => TextExtractor.Extract(new byte[] { })
      );
    }

    [TestMethod]
    public void TestExtractMethodWithMissingFile()
    {
      Assert.ThrowsException<FileNotFoundException>(
        () => TextExtractor.Extract(string.Empty)
      );
    }

    [TestMethod]
    [DeploymentItem(@"Documents\Empty.docx")]
    public void TestExtractMethodWithEmptyFile()
    {
      Assert.ThrowsException<NotSupportedException>(
        () => TextExtractor.Extract(@"Documents\Empty.docx")
      );
    }

    [TestMethod]
    [DeploymentItem(@"Documents\Invalid.docx")]
    public void TestExtractMethodWithInvalidFile()
    {
      Assert.ThrowsException<NotSupportedException>(
        () => TextExtractor.Extract(@"Documents\Invalid.docx")
      );
    }

    [TestMethod]
    [DeploymentItem(@"Documents\Valid.docx")]
    public void TestExtractMethodWithValidFile()
    {
      var text = TextExtractor.Extract(@"Documents\Valid.docx");

      Assert.IsTrue(text == "This is a Word document.");
    }

    [TestMethod]
    [DeploymentItem(@"Documents\ValidWithWhitespaceNoPreserve.docx")]
    public void TestExtractMethodWithValidFileWithWhitespaceNoPreserveFile()
    {
      var text = TextExtractor.Extract(@"Documents\ValidWithWhitespaceNoPreserve.docx");

      Assert.IsTrue(text == "This is a Word document.");
    }

    [TestMethod]
    [DeploymentItem(@"Documents\ValidWithWhitespacePreserve.docx")]
    public void TestExtractMethodWithValidFileWithWhitespacePreserveFile()
    {
      var text = TextExtractor.Extract(@"Documents\ValidWithWhitespacePreserve.docx");

      Assert.IsTrue(text == "This is a Word document.");
    }

    [TestMethod]
    [DeploymentItem(@"Documents\ValidWithFormatting.docx")]
    public void TestExtractMethodWithValidFileWithFormatting()
    {
      var actualText = TextExtractor.Extract(@"Documents\ValidWithFormatting.docx");

      var expectedText =
        string.Concat(
          "This is a Word document. It spans over multiple paragraphs with line breaks. ",
          "It also has some basic formatting."
        );

      Assert.IsTrue(actualText == expectedText);
    }

    [TestMethod]
    [DeploymentItem(@"Documents\ValidWithTables.docx")]
    public void TestExtractMethodWithValidWithTables()
    {
      var actualText = TextExtractor.Extract(@"Documents\ValidWithTables.docx");

      var expectedText =
        string.Concat(
          "Heading 1 This is a document. ",
          "Column 1 Column 2 Column 3 Cell 1x1 Cell 1x2 Cell 1x3 Cell 2x1 Cell 2x2 Cell 2x3"
        );

      Assert.IsTrue(actualText == expectedText);
    }

    [TestMethod]
    public void TestIsValidFileTypeMethodWithUndefinedFile()
    {
      Assert.ThrowsException<ArgumentNullException>(
        () => TextExtractor.IsValidFileType((string)null)
      );
    }

    [TestMethod]
    public void TestIsValidFileTypeMethodWithUndefinedStream()
    {
      Assert.ThrowsException<ArgumentNullException>(
        () => TextExtractor.IsValidFileType((Stream)null)
      );
    }

    [TestMethod]
    public void TestIsValidFileTypeMethodWithUndefinedBytes()
    {
      Assert.ThrowsException<ArgumentNullException>(
        () => TextExtractor.IsValidFileType((byte[])null)
      );
    }

    [TestMethod]
    public void TestIsValidFileTypeMethodWithMissingStream()
    {
      Assert.IsFalse(TextExtractor.IsValidFileType(new MemoryStream()));
    }

    [TestMethod]
    public void TestIsValidFileTypeMethodWithMissingBytes()
    {
      Assert.IsFalse(TextExtractor.IsValidFileType(new byte[] { }));
    }

    [TestMethod]
    public void TestIsValidFileTypeMethodWithMissingFile()
    {
      Assert.ThrowsException<FileNotFoundException>(
        () => TextExtractor.IsValidFileType(string.Empty)
      );
    }

    [TestMethod]
    [DeploymentItem(@"Documents\Empty.docx")]
    public void TestIsValidFileTypeMethodWithEmptyFile()
    {
      Assert.IsFalse(TextExtractor.IsValidFileType(@"Documents\Empty.docx"));
    }

    [TestMethod]
    [DeploymentItem(@"Documents\Invalid.docx")]
    public void TestIsValidFileTypeMethodWithInvalidFile()
    {
      Assert.IsFalse(TextExtractor.IsValidFileType(@"Documents\Invalid.docx"));
    }

    [TestMethod]
    [DeploymentItem(@"Documents\Valid.docx")]
    public void TestIsValidFileTypeMethodWithValidFile()
    {
      Assert.IsTrue(TextExtractor.IsValidFileType(@"Documents\Valid.docx"));
    }
  }
}