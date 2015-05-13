using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Mery.DotNetLib;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.ComponentModel;
using System.IO;
using System.Linq;


namespace TakamiChie.Mery.AutoLoader
{

  public class DllInstance : DllInstanceBase
  {
    const int WM_USER = 0x0400;
    const int ME_FIRST = WM_USER + 0x0400;
    const int ME_OUTPUT_STRING = ME_FIRST + 48;
    
    const int FLAG_OPEN_OUTPUT = 1;
    const int FLAG_CLOSE_OUTPUT = 2;
    const int FLAG_FOCUS_OUTPUT = 4;
    const int FLAG_CLEAR_OUTPUT = 8;

    [DllImport("User32.dll", EntryPoint = "SendMessageW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, int wParam, String lParam);

    #region 標準処理
    //=====================================================================
    // 標準処理

    ///// <summary>
    ///// メニュー実行．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    public void OnCommand(IntPtr hWnd)
    {
      string n;
      GetName(out n);
      MessageBox.Show("現在設定はありません", n);
    }

    ///// <summary>
    ///// メニュー状態取得．
    ///// このメソッドが未定義で OnCommand が定義済みの場合は「実行可能」とする．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    ///// <param name="isChecked">押し込み状態</param>
    ///// <returns>実行可能な場合は true</returns>
    //public bool QueryStatus(IntPtr hWnd, out bool isChecked) {
    //    return true;
    //}

    ///// <summary>
    ///// プラグイン名．
    ///// 未定義の場合は実行ファイル名とする．
    ///// </summary>
    ///// <param name="name">プラグイン名</param>
    public void GetName(out String name)
    {
      name = "Mery_AutoLoaderPlugin";
    }

    ///// <summary>
    ///// プラグインバージョン．
    ///// 未定義の場合は非表示とする．
    ///// </summary>
    ///// <param name="version">プラグインバージョン</param>
    public void GetVersion(out String version)
    {
      var ver =
          System.Diagnostics.FileVersionInfo.GetVersionInfo(
          System.Reflection.Assembly.GetExecutingAssembly().Location);
      version = "ver " + ver;
    }

    ///// <summary>
    ///// プロパティ画面が起動可能かの問い合わせ．
    ///// このメソッドが未定義で SetProperty が定義されている場合は「起動可能」とする．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    ///// <returns>起動可能な場合は true</returns>
    //public bool QueryProperty(IntPtr hWnd) {
    //    return false;
    //}

    ///// <summary>
    ///// プロパティ画面起動
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    //public void SetProperty(IntPtr hWnd) {
    //}
    #endregion


    #region DLL 初期化・終了
    //=====================================================================
    // DLL 初期化・終了

    ///// <summary>
    ///// DLL 初期化処理(もどき)．
    ///// </summary>
    ///// <param name="hModule">ラッパー DLL のモジュールハンドル</param>
    //public void OnDllAttach(IntPtr hModule) {
    //}

    ///// <summary>
    ///// DLL 終了処理(もどき)．
    ///// </summary>
    ///// <param name="hModule">ラッパー DLL のモジュールハンドル</param>
    //public void OnDllDetach(IntPtr hModule) {
    //}
    #endregion


    #region OnEvents のディスパッチ
    //=====================================================================
    // OnEvents のディスパッチ

    ///// <summary>
    ///// イベントに対する処理を記述．
    ///// 個別の On*** を記述していた場合はそちらが優先して呼び出される．
    ///// これを記述した場合は常にアイドルイベントが呼ばれるため，できれば個別に実装するべし．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    ///// <param name="uEvent">イベントコード</param>
    ///// <param name="lParam">パラメータ</param>
    //public void OnEvents(IntPtr hWnd, UInt32 uEvent, IntPtr lParam) {
    //}

    ///// <summary>
    ///// エディタを起動した時．
    ///// </summary>
    ///// <param name="hWnd">Mery の標準プロセス</param>
    ///// <param name="lParam">？？？</param>
    //public void OnCreate(IntPtr hWnd, IntPtr lParam) {
    //}

    ///// <summary>
    ///// エディタを終了した時．
    ///// </summary>
    ///// <param name="hWnd">Mery の標準プロセス</param>
    //public void OnClose(IntPtr hWnd) {
    //}

    ///// <summary>
    ///// フレームが作成された時．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    //public void OnCreateFrame(IntPtr hWnd) {
    //}

    ///// <summary>
    ///// フレームが破棄された時．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    //public void OnCloseFrame(IntPtr hWnd) {
    //}

    ///// <summary>
    ///// フォーカスを取得した時．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    //public void OnSetFocus(IntPtr hWnd) {
    //}

    ///// <summary>
    ///// フォーカスを失った時．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    //public void OnKillFocus(IntPtr hWnd) {
    //}

    ///// <summary>
    ///// ファイルを開いた時．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    //public void OnFileOpened(IntPtr hWnd) {
    //}

    ///// <summary>
    ///// ファイルを保存した時．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    public void OnFileSaved(IntPtr hWnd)
    {
      var manager = new CommandManager(new Editor(hWnd));
      if (manager.hasKey("run"))
      {
        // プロセス起動準備
        var procinfo = new ProcessStartInfo(manager.getPath("run"));
        procinfo.Arguments = manager.hasKey("args") ? manager.getExtractVar("args") : "";
        procinfo.CreateNoWindow = true;
        procinfo.RedirectStandardOutput = true;
        procinfo.RedirectStandardError = true;
        procinfo.UseShellExecute = false;
        var stdout = "";
        var stderr = "";
        var usecon = false;
        try
        {
          // 起動
          var proc = Process.Start(procinfo);
          stdout = proc.StandardOutput.ReadToEnd();
          stderr = proc.StandardError.ReadToEnd();

          // 結果表示
          switch (manager.hasKey("out") ? manager.get("out") : "console")
          {
            case "file":
              // ファイル出力
              if (manager.hasKey("outfile"))
              {
                using (var stream = new StreamWriter(manager.getExtractVar("outfile")))
                {
                  stream.Write(stdout);
                }
              }
              else
              {
                throw new InvalidDataException("outfileコマンドが見つかりません");
              }
              break;
            case "console":
              // コンソール出力
              if (!usecon)
              {
                SendMessage(hWnd, ME_OUTPUT_STRING, FLAG_CLEAR_OUTPUT, "");
                usecon = true;
              }
              SendMessage(hWnd, ME_OUTPUT_STRING, 0, stdout);
              break;
          }
        }
        catch (Exception e)
        {
          stderr = e.Message;
        }
        if (stderr != "")
        {
          MessageBox.Show(stderr, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        // メッセージ表示
        if (manager.hasKey("console"))
        {
          if (!usecon)
          {
            SendMessage(hWnd, ME_OUTPUT_STRING, FLAG_CLEAR_OUTPUT, "");
            usecon = true;
          }
          SendMessage(hWnd, ME_OUTPUT_STRING, 0, manager.getExtractVar("console"));
        }

        if (manager.hasKey("dialog"))
        {
          MessageBox.Show(manager.getExtractVar("dialog"));
        }

      }
    }

    ///// <summary>
    ///// 更新の状態が変更された時．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    //public void OnModified(IntPtr hWnd) {
    //}

    ///// <summary>
    ///// カーソルが移動した時．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    //public void OnCaretMoved(IntPtr hWnd) 
    //}

    ///// <summary>
    ///// スクロールされた時．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    //public void OnScroll(IntPtr hWnd) {
    //}

    ///// <summary>
    ///// 選択範囲が変更された時．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    //public void OnSelChanged(IntPtr hWnd) {
    //}

    ///// <summary>
    ///// テキストが変更された時．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    //public void OnChanged(IntPtr hWnd) {
    //}

    ///// <summary>
    ///// 文字が入力された時．
    ///// ASCII の場合はその文字が渡される．
    ///// IME 経由の場合文字単位でイベントが呼ばれ，その後文字数分だけ input=0 が渡される．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    ///// <param name="ch">入力された文字</param>
    //public void OnChar(IntPtr hWnd, Char ch) {
    //}

    ///// <summary>
    ///// 編集モードが変更された時．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    //public void OnModeChanged(IntPtr hWnd) {
    //}

    ///// <summary>
    ///// アクティブな文書が変更された時．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    //public void OnDocumentSelectChanged(IntPtr hWnd) {
    //}

    ///// <summary>
    ///// 文書を閉じた時．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    //public void OnDocumentClose(IntPtr hWnd) {
    //}

    ///// <summary>
    ///// タブが移動された時．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    //public void OnTabMoved(IntPtr hWnd) {
    //}

    ///// <summary>
    ///// カスタムバーを閉じようとしたとき．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    ///// <param name="id">カスタムバーID</param>
    ///// <param name="pos">カスタムバー位置</param>
    ///// <param name="flags">フラグ</param>
    //public void OnCustomBarClosing(IntPtr hWnd, int id, int pos, UInt32 flags) {
    //}

    ///// <summary>
    ///// カスタムバーを閉じたとき．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    ///// <param name="id">カスタムバーID</param>
    ///// <param name="pos">カスタムバー位置</param>
    ///// <param name="flags">フラグ</param>
    //public void OnCustomBarClosed(IntPtr hWnd, int id, int pos, UInt32 flags) {
    //}

    ///// <summary>
    ///// ツールバーを閉じた時．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    //public void OnToolBarClosed(IntPtr hWnd) {
    //}

    ///// <summary>
    ///// ツールバーが表示された時．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    //public void OnToolBarShow(IntPtr hWnd) {
    //}

    ///// <summary>
    ///// アイドル時．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    //public void OnIdle(IntPtr hWnd) {
    //}
    #endregion


    #region PreTranslateMessage のディスパッチ
    ///// <summary>
    ///// PreTranslateMessage 処理．
    ///// ただし全てのメッセージが送られるわけではなく，Mery が渡してくれる範囲のみ．
    ///// 頻繁に呼ばれるため，個別定義可能な処理は個別定義にするべし．
    ///// </summary>
    ///// <param name="hEditorWnd">対象のエディタハンドル</param>
    ///// <param name="hWnd">PreTranslateMessage の対象ハンドル</param>
    ///// <param name="message">メッセージコード</param>
    ///// <param name="wParam">WPARAM 引数</param>
    ///// <param name="lParam">LPARAM 引数</param>
    ///// <param name="time">メッセージの時間</param>
    ///// <param name="x">マウス X 座標</param>
    ///// <param name="y">マウス Y 座標</param>
    ///// <returns>メッセージ処理を継続する場合は false．</returns>
    //public bool PreTranslateMessage(IntPtr hEditorWnd, IntPtr hWnd, uint message, IntPtr wParam, IntPtr lParam, Int32 time, int x, int y) {
    //    return false;
    //}

    ///// <summary>
    ///// マウスを動かした時に呼ばれます．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    ///// <param name="key">各種の仮想キーが押されているかどうかを示します．各ビットの意味は <see cref="KEY_STATE"/> に定義されています．</param>
    ///// <param name="x">X 座標</param>
    ///// <param name="y">Y 座標</param>
    ///// <returns>メッセージ処理を継続する場合は false．</returns>
    //public bool OnMouseMove(IntPtr hWnd, int key, int x, int y) {
    //    return false;
    //}

    ///// <summary>
    ///// マウス左ボタンを押した時に呼ばれます．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    ///// <param name="key">各種の仮想キーが押されているかどうかを示します．各ビットの意味は <see cref="KEY_STATE"/> に定義されています．</param>
    ///// <param name="x">X 座標</param>
    ///// <param name="y">Y 座標</param>
    ///// <returns>メッセージ処理を継続する場合は false．</returns>
    //public bool OnLButtonDown(IntPtr hWnd, int key, int x, int y) {
    //    return false;
    //}

    ///// <summary>
    ///// マウス左ボタンを離した時に呼ばれます．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    ///// <param name="key">各種の仮想キーが押されているかどうかを示します．各ビットの意味は <see cref="KEY_STATE"/> に定義されています．</param>
    ///// <param name="x">X 座標</param>
    ///// <param name="y">Y 座標</param>
    ///// <returns>メッセージ処理を継続する場合は false．</returns>
    //public bool OnLButtonUp(IntPtr hWnd, int key, int x, int y) {
    //    return false;
    //}

    ///// <summary>
    ///// マウス左ボタンをダブルクリックした時に呼ばれます．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    ///// <param name="key">各種の仮想キーが押されているかどうかを示します．各ビットの意味は <see cref="KEY_STATE"/> に定義されています．</param>
    ///// <param name="x">X 座標</param>
    ///// <param name="y">Y 座標</param>
    ///// <returns>メッセージ処理を継続する場合は false．</returns>
    //public bool OnLButtonDblClk(IntPtr hWnd, int key, int x, int y) {
    //    return false;
    //}

    ///// <summary>
    ///// マウス右ボタンを押した時に呼ばれます．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    ///// <param name="key">各種の仮想キーが押されているかどうかを示します．各ビットの意味は <see cref="KEY_STATE"/> に定義されています．</param>
    ///// <param name="x">X 座標</param>
    ///// <param name="y">Y 座標</param>
    ///// <returns>メッセージ処理を継続する場合は false．</returns>
    //public bool OnRButtonDown(IntPtr hWnd, int key, int x, int y) {
    //    return false;
    //}

    ///// <summary>
    ///// マウス右ボタンを離した時に呼ばれます．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    ///// <param name="key">各種の仮想キーが押されているかどうかを示します．各ビットの意味は <see cref="KEY_STATE"/> に定義されています．</param>
    ///// <param name="x">X 座標</param>
    ///// <param name="y">Y 座標</param>
    ///// <returns>メッセージ処理を継続する場合は false．</returns>
    //public bool OnRButtonUp(IntPtr hWnd, int key, int x, int y) {
    //    return false;
    //}

    ///// <summary>
    ///// マウス右ボタンをダブルクリックした時に呼ばれます．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    ///// <param name="key">各種の仮想キーが押されているかどうかを示します．各ビットの意味は <see cref="KEY_STATE"/> に定義されています．</param>
    ///// <param name="x">X 座標</param>
    ///// <param name="y">Y 座標</param>
    ///// <returns>メッセージ処理を継続する場合は false．</returns>
    //public bool OnRButtonDblClk(IntPtr hWnd, int key, int x, int y) {
    //    return false;
    //}

    ///// <summary>
    ///// マウスホイールを動かしたときに呼ばれます．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    ///// <param name="key">各種の仮想キーが押されているかどうかを示します．各ビットの意味は <see cref="KEY_STATE"/> に定義されています．</param>
    ///// <param name="delta">移動量</param>
    ///// <param name="x">X 座標</param>
    ///// <param name="y">Y 座標</param>
    ///// <returns>メッセージ処理を継続する場合は false．</returns>
    //public bool OnMouseWheel(IntPtr hWnd, int key, int delta, int x, int y) {
    //    return false;
    //}

    ///// <summary>
    ///// キーを押したときに呼ばれます．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    ///// <param name="keycode">仮想キーコード．<see cref="VIRTUAL_KEY"/> に定義されています．</param>
    ///// <param name="repeat">リピートカウント</param>
    ///// <param name="previous">直前のキー状態が指定されます．true の場合，メッセージが送られる前からキーが押されています．</param>
    ///// <returns>メッセージ処理を継続する場合は false．</returns>
    //public bool OnKeyDown(IntPtr hWnd, int keycode, int repeat, bool previous) {
    //    return false;
    //}

    ///// <summary>
    ///// キーを離したときに呼ばれます．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    ///// <param name="keycode">仮想キーコード．<see cref="VIRTUAL_KEY"/> に定義されています．</param>
    ///// <returns>メッセージ処理を継続する場合は false．</returns>
    //public bool OnKeyUp(IntPtr hWnd, int keycode) {
    //    return false;
    //}

    ///// <summary>
    ///// 文字入力時に呼ばれます．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    ///// <param name="ch">入力文字</param>
    ///// <param name="repeat">リピートカウント</param>
    ///// <param name="alt">ALT キーが押されているか</param>
    ///// <param name="previous">直前のキー状態が指定されます．true の場合，メッセージが送られる前からキーが押されています．</param>
    ///// <param name="convert">変換状態が指定されます．</param>
    ///// <returns>メッセージ処理を継続する場合は false．</returns>
    //public bool OnChar(IntPtr hWnd, Char ch, int repeat, bool alt, bool previous, bool convert) {
    //    return false;
    //}

    ///// <summary>
    ///// IME 経由の文字列入力時に呼ばれます．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    ///// <param name="input">入力文字列</param>
    ///// <returns>メッセージ処理を継続する場合は false．</returns>
    //public bool OnImeChar(IntPtr hWnd, string input) {
    //    return false;
    //}

    ///// <summary>
    ///// システムキーを押したときに呼ばれます．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    ///// <param name="keycode">仮想キーコード．<see cref="VIRTUAL_KEY"/> に定義されています．</param>
    ///// <param name="repeat">リピートカウント</param>
    ///// <param name="alt">ALT キーが押されているか</param>
    ///// <param name="previous">直前のキー状態が指定されます．true の場合，メッセージが送られる前からキーが押されています．</param>
    ///// <returns>メッセージ処理を継続する場合は false．</returns>
    //public bool OnSysKeyDown(IntPtr hWnd, int keycode, int repeat, bool alt, bool previous) {
    //    return false;
    //}

    ///// <summary>
    ///// システムキーを離したときに呼ばれます．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    ///// <param name="keycode">仮想キーコード．<see cref="VIRTUAL_KEY"/> に定義されています．</param>
    ///// <param name="alt">ALT キーが押されているか</param>
    ///// <returns>メッセージ処理を継続する場合は false．</returns>
    //public bool OnSysKeyUp(IntPtr hWnd, int keycode, bool alt) {
    //    return false;
    //}

    ///// <summary>
    ///// タイマ処理．
    ///// <see cref="SetTimer"/> でタイマを作成する．
    ///// <see cref="Mery.DotNetLib.DllInstanceBase"/> の継承が必要．
    ///// </summary>
    ///// <param name="hWnd">対象のエディタハンドル</param>
    ///// <param name="id"><see cref="SetTimer"/> で作成したタイマ ID</param>
    //public void OnTimer(IntPtr hWnd, UInt16 id) {
    //}

    #endregion

  }
}
