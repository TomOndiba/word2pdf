// WordConverter.cs
// http://blog.jhashimoto.jp/entry/20120604/1338801745

using MSWord = Microsoft.Office.Interop.Word;

namespace WordToPdf
{
  /// <summary>
  /// Wordファイルを変換する機能を提供します。
  /// </summary>
  public static class WordConverter
  {
    /// <summary>
    /// WordファイルをPDFとして保存します。
    /// </summary>
    /// <param name="wordFilePathName">Wordファイルのパス付きファイル名。</param>
    /// <param name="saveAsPathName">保存するPDFのパス付きファイル名。</param>
    /// <remarks>
    /// <para>
    /// Word 2013がインストールされている必要があります。
    /// </para>
    /// </remarks>
    public static void SaveAsPdf(string wordFilePathName, string saveAsPathName)
    {
      // Word 2013がインストールされている必要があります。
      MSWord.ApplicationClass application = null;
      MSWord.Documents documents = null;
      MSWord.DocumentClass document = null;

      // refキーワードと共に渡すので、変数である必要がある。
      object missing = System.Type.Missing;

      try
      {
        application = new MSWord.ApplicationClass();
        /*
         * application.Documents.Open(...は、Documentsオブジェクトの解放処理ができないので不可。
         * 必ず変数経由でComRelease.FinalReleaseComObjectsを呼び出すこと。
         */
        documents = application.Documents;

        object filePathName = wordFilePathName;
        document = (MSWord.DocumentClass)documents.Open(
            ref filePathName, ref missing, ref missing, ref missing, ref missing
            , ref missing, ref missing, ref missing, ref missing, ref missing
            , ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

        // ExportAsFixedFormatのパラメータは以下を参照
        // http://msdn.microsoft.com/ja-jp/library/microsoft.office.tools.word.document.exportasfixedformat.aspx
        // http://msdn.microsoft.com/en-us/library/bb412305.aspx
        document.ExportAsFixedFormat(
            saveAsPathName,
            MSWord.WdExportFormat.wdExportFormatPDF,
            false,
            MSWord.WdExportOptimizeFor.wdExportOptimizeForPrint,
            MSWord.WdExportRange.wdExportAllDocument,
            0,
            0,
            MSWord.WdExportItem.wdExportDocumentWithMarkup,
            true,
            true,
            MSWord.WdExportCreateBookmarks.wdExportCreateWordBookmarks,
            true,
            true,
            false,
            ref missing);
      }
      finally
      {
        if (document != null)
        {
          try
          {
            document.Close(ref missing, ref missing, ref missing);
          }
          catch { }
        }
        if (application != null)
        {
          try
          {
            application.Quit(ref missing, ref missing, ref missing);
          }
          catch { }
        }
        Com.ComRelease.FinalReleaseComObjects(document, documents, application);
      }
    }
  }
}
