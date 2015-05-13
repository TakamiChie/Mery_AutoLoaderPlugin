using Mery.DotNetLib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace TakamiChie.Mery.AutoLoader
{
  class CommandManager
  {
    private Dictionary<String, String> map;
    private Dictionary<String, String> pathMap;

    public CommandManager(Editor editor)
    {
      this.map = new Dictionary<string, string>();
      this.pathMap = new Dictionary<string, string>();
      pathMapping(editor);
      codeMapping(editor);
    }

    /// <summary>
    /// 値がコマンドリストに存在するかどうか
    /// </summary>
    /// <param name="key">キー</param>
    /// <returns>キーが存在すればtrue</returns>
    public bool hasKey(String key)
    {
      return map.ContainsKey(key);
    }

    /// <summary>
    /// コマンドリストの値を取得する
    /// </summary>
    /// <param name="key">キー</param>
    /// <returns>キー内の値</returns>
    public string get(String key)
    {
      return map[key];
    }

    /// <summary>
    /// コマンドリストの値を取得。実行ファイルが見つかればそのパスを取得する
    /// </summary>
    /// <param name="key">キー</param>
    /// <returns>キー内の値</returns>
    public string getPath(String key)
    {
      return which(get(key));
    }

    /// <summary>
    /// コマンドリストの値を取得。パス変数を展開する
    /// </summary>
    /// <param name="key">キー</param>
    /// <returns>キー内の値</returns>
    public string getExtractVar(String key)
    {
      var s = get(key);
      foreach (var item in pathMap)
      {
        s = s.Replace(item.Key, item.Value);
      }
      return s;
    }

    /// <summary>
    /// パス変数の一覧を生成する
    /// </summary>
    /// <param name="editor">エディタオブジェクト</param>
    private void pathMapping(Editor editor)
    {
      var path = editor.GetDocumentPath();
      this.pathMap.Add("$1", path);
      this.pathMap.Add("$dn1", Path.GetDirectoryName(path));
      this.pathMap.Add("$fn1", Path.GetFileName(path));
      this.pathMap.Add("$ex1", Path.GetExtension(path));
      this.pathMap.Add("$fnwe1", Path.GetFileNameWithoutExtension(path));
    }

    /// <summary>
    /// コード変数の一覧を生成する
    /// </summary>
    /// <param name="editor">エディタオブジェクト</param>
    private void codeMapping(Editor editor)
    {
      var lc = editor.GetLines(true);
      var regex = new Regex(@".*\*{2}\s*(\w+):(.*)");
      for (int i = 0; i < lc - 1; i++)
      {
        var s = editor.GetLine(i, true);
        var m = regex.Match(s);
        if (m.Success)
        {
          this.map.Add(m.Groups[1].Value, m.Groups[2].Value);
        }
        else
        {
          break;
        }
      }
    }

    /// <summary>
    /// ファイル名をパス名に変換する
    /// </summary>
    /// <param name="fileName">ファイル名</param>
    /// <returns>パス</returns>
    private string which(String fileName)
    {
      return Environment.GetEnvironmentVariable("PATH").Split(';')
        .Select(x => Path.Combine(x, fileName))
        .SelectMany(_ => Environment.GetEnvironmentVariable("PATHEXT").Split(';')
          .Concat(new String[] { Path.GetExtension(fileName) }),
            (p, ext) => Path.ChangeExtension(p, ext))
        .Where(File.Exists).FirstOrDefault();
    }
  }
}
