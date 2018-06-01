using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace MintyPeterson.TextExtractor.Tests
{
  [TestClass]
  public class TextExtractorUnitTest
  {
    [TestMethod]
    public void TextExtractMethodWithUndefinedFile()
    {
      Assert.ThrowsException<ArgumentNullException>(
        () => TextExtractor.Extract((string)null)
      );
    }

    [TestMethod]
    public void TextExtractMethodWithUndefinedStream()
    {
      Assert.ThrowsException<ArgumentNullException>(
        () => TextExtractor.Extract((Stream)null)
      );
    }

    [TestMethod]
    public void TextExtractMethodWithUndefinedBytes()
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
      var text = TextExtractor.Extract(@"Documents\ValidWithFormatting.docx");

      Assert.IsTrue(text == "This is a Word document. It spans over multiple paragraphs with line breaks. It also has some basic formatting.");
    }

    [TestMethod]
    [DeploymentItem(@"Documents\ValidWithTables.docx")]
    public void TestExtractMethodWithValidWithTables()
    {
      var text = TextExtractor.Extract(@"Documents\ValidWithTables.docx");

      Assert.IsTrue(text == "Heading 1 This is a document. Column 1 Column 2 Column 3 Cell 1x1 Cell 1x2 Cell 1x3 Cell 2x1 Cell 2x2 Cell 2x3");
    }
  }
}
