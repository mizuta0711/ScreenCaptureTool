using ScreenCaptureTool.Models.CaptureItem;

using System.Drawing;

namespace ScreenCaptureTool.Controllers.CaptureProvider
{
    /// <summary>
    /// キャプチャー機能の基底クラス
    /// </summary>
    public abstract class CaptureProvider
    {
        #region Property

        /// <summary>
        /// キャプチャーアイテム
        /// </summary>
        public CaptureItemBase CaptureItem { protected set; get; }

        #endregion Property

        #region Constructor

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="item">キャプチャーアイテム</param>
        public CaptureProvider(CaptureItemBase item)
        {
            CaptureItem = item;
        }

        #endregion Constructor

        #region Methods(Public)

        /// <summary>
        /// キャプチャー実行
        /// </summary>
        /// <returns>画像(失敗時はnull)</returns>
        public virtual Bitmap? Capture()
        {
            return CaptureItem?.Capture();
        }

        #endregion Methods(Public)
    }
}