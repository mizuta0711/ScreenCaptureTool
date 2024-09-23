using System;
using System.Collections.ObjectModel;

namespace ScreenCaptureTool
{
    /// <summary>
    /// アプリケーションの設定を保持するクラス
    /// </summary>
    [Serializable]
    public class AppSettings
    {
        public enum CaptureType
        {
            ScreenRect,     // 画面の矩形
            Window          // ウィンドウ
        }

        /// <summary>
        /// キャプチャー範囲：X
        /// </summary>
        public int CaptureLeft { get; set; }

        /// <summary>
        /// キャプチャー範囲：Y
        /// </summary>
        public int CaptureTop { get; set; }

        /// <summary>
        /// キャプチャー範囲：幅
        /// </summary>
        public int CaptureWidth { get; set; }

        /// <summary>
        /// キャプチャー範囲：高さ
        /// </summary>
        public int CaptureHeight { get; set; }

        /// <summary>
        /// ウィンドウ位置：X
        /// </summary>
        public double WindowLeft { get; set; }

        /// <summary>
        /// ウィンドウ位置：Y
        /// </summary>
        public double WindowTop { get; set; }

        /// <summary>
        /// ウィンドウサイズ：幅
        /// </summary>
        public double WindowWidth { get; set; }

        /// <summary>
        /// ウィンドウサイズ：高さ
        /// </summary>
        public double WindowHeight { get; set; }

        /// <summary>
        /// キャプチャーするウィンドウタイトル
        /// </summary>
        public string CaptureWindowTitle { get; set; }

        /// <summary>
        /// チャプチャータイプ
        /// </summary>
        public CaptureType SelectedCaptureType { get; set; }

        /// <summary>
        /// サムネイル画像サイズ
        /// </summary>
        public int ThumbnailSize { get; set; } = 200;  // デフォルトサイズ

        /// <summary>
        /// 保存ファイル名の一覧
        /// </summary>
        public ObservableCollection<string> SaveFileNames = new ObservableCollection<string>();

        /// <summary>
        /// 画像保存先フォルダ
        /// </summary>
        public string SaveFolderPath { get; set; }
    }
}