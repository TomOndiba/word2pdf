// Program.cs
// 
// ＜概要＞
// コマンドライン上でWord->PDFの変換を行います．
// COM(Component Object Model)経由でWord2013の
// 機能にアクセスしています．
// 
// ＜使い方＞
// wordFilePathName: 入力のWordファイルパス(*.doc,*.docx)
// saveAsPathName: PDFファイルパスの出力先(*.pdf)
// > word2pdf wordFilePathName saveAsPathName
// 
// ＜動作環境＞
// ・Microsoft Word 2013
// ・Microsoft .NET Framework 4.5
// 
// ＜バージョン履歴＞
// Ver 1.0.1  2014/09/05
// Ver 1.1.0  2015/12/13 常に読み取り専用かつVisible=falseで開くよう修正
// 
// © 2014 saito3

/// <summary>
/// Convert MS Word DOC to PDF
/// </summary>

namespace WordToPdf
{
  class Program
  {
    static int Main(string[] args)
    {
      if (args.Length != 2)
      {
        System.Console.WriteLine("Usage: .exe wordFilePathName(*.doc,*.docx) saveAsPathName(*.pdf)");
        return 1;
      }
      string wordFilePathName = System.IO.Path.GetFullPath(args[0]);
      string saveAsPathName = System.IO.Path.GetFullPath(args[1]);
      try
      {
        WordConverter.SaveAsPdf(wordFilePathName, saveAsPathName);
      }
      catch (System.Exception e)
      {
        System.Console.WriteLine(e);
        return 1;
      }
      return 0;
    }
  }
}
