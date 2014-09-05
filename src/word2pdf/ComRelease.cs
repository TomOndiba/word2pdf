// ComRelease.cs
// http://blog.jhashimoto.jp/entry/20120604/1338801745

/// <summary>
/// 複数のCOMオブジェクトの参照カウントを０までデクリメントし、解放します。
/// </summary>

using System.Runtime.InteropServices;

namespace Com
{
  /// <summary>
  /// COMオブジェクトを解放する機能を提供します。
  /// </summary>
  public static class ComRelease
  {
    /// <summary>
    /// 複数のCOMオブジェクトの参照カウントを０までデクリメントし、解放します。
    /// </summary>
    /// <param name="objects">解放するCOMオブジェクトの配列。</param>
    /// <remarks>解放は配列の要素順に行います。</remarks>
    public static void FinalReleaseComObjects(params object[] objects)
    {
      foreach (object o in objects)
      {
        try
        {
          if (o == null || Marshal.IsComObject(o) == false)
          {
            continue;
          }
          Marshal.FinalReleaseComObject(o);
        }
        catch (System.Exception) { }
      }
    }
  }
}
