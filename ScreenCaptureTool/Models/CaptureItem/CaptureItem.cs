using System;
using System.Drawing;
using System.IO;

namespace ScreenCaptureTool.Models.CaptureItem
{
    /// <summary>
    /// キャプチャーアイテム
    /// </summary>
    [Serializable]
    internal abstract class CaptureItem
    {
        #region Properties

        /// <summary>
        /// サブフォルダ名
        /// </summary>
        public string? SubFolder { get; protected set; }

        /// <summary>
        /// ファイル名（拡張子なし）
        /// </summary>
        public string FileName { get; protected set; }

        /// <summary>
        /// 拡張子
        /// </summary>
        protected readonly string FileExtension = ".png";

        #endregion Properties

        #region Constructor

        /// <summary>
        /// コンストラクタ
        /// </summary>
        protected CaptureItem()
        {
            SubFolder = null;
            FileName = DateTime.Now.ToString("yyyyMMddHHmmss");
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="fileName">ファイル名(拡張子なし)</param>
        /// <param name="subFilder">サブフォルダ名</param>
        protected CaptureItem(string fileName, string? subFilder = null)
        {
            FileName = fileName;
            SubFolder = subFilder;
        }

        #endregion Constructor

        #region Methods

        #region Methods(Public)

        /// <summary>
        /// アイテムのパスをフルパスで取得する
        /// </summary>
        /// <param name="rootFolder">親フォルダ</param>
        /// <returns>フルパス</returns>
        public string GetFilePath(string rootFolder)
        {
            string fullPath = rootFolder;
            if (String.IsNullOrEmpty(SubFolder) == false)
            {
                fullPath = Path.Combine(fullPath, SubFolder);
            }
            return Path.ChangeExtension(Path.Combine(fullPath, FileName), FileExtension);
        }

        #endregion Methods(Public)

        #region Methods(Abstract)

        /// <summary>
        /// キャプチャーを行う
        /// </summary>
        /// <returns>画像(失敗時はnull)</returns>
        public abstract Bitmap? Capture();

        #endregion Methods(Abstract)

        #endregion Methods
    }
}