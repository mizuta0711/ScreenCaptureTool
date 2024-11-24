using ScreenCaptureTool.Models.CaptureItem;
using ScreenCaptureTool.Windows;

using System;
using System.Drawing;

namespace ScreenCaptureTool.Controllers.CaptureProvider
{
    /// <summary>
    /// ユーザー指定の矩形範囲キャプチャー機能
    /// </summary>
    public class ManualScreenRectCaptureProvider : CaptureProvider
    {
        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="item">キャプチャーアイテム</param>
        public ManualScreenRectCaptureProvider(CaptureItemBase item) : base(item)
        {
            // キャプチャーアイテムの型チェック
            if (!(item is ManualScreenRectCaptureItem))
            {
                throw new System.ArgumentException("item must be ManualScreenRectCaptureItem", nameof(item));
            }
        }

        /// <summary>
        /// キャプチャーを行う
        /// </summary>
        /// <returns>画像(失敗時はnull)</returns>
        /// <exception cref="InvalidOperationException">キャプチャーアイテムがManualScreenRectCaptureItemでない場合</exception>
        public override Bitmap? Capture()
        {
            var overlayWindow = new CaptureRectOverlayWindow();

            if (CaptureItem is ManualScreenRectCaptureItem manualScreenRectCaptureItem)
            {
                if (manualScreenRectCaptureItem.LocationFixed && manualScreenRectCaptureItem.SizeFixed)
                {
                    // 位置もサイズも固定されている場合はそのままキャプチャー
                    return base.Capture();
                }

                // 範囲情報を設定する
                overlayWindow.SetCaptureItem(manualScreenRectCaptureItem);

                // オーバーレイウィンドウを表示
                if (overlayWindow.ShowDialog() == true)
                {
                    // 選択された矩形が空の場合はキャンセル
                    if (overlayWindow.SelectedRect.IsEmpty)
                    {
                        return null;
                    }

                    // 選択された矩形を設定
                    manualScreenRectCaptureItem.TargetRect = new Rectangle(
                        (int)Math.Floor(overlayWindow.SelectedRect.X),
                        (int)Math.Floor(overlayWindow.SelectedRect.Y),
                        (int)Math.Floor(overlayWindow.SelectedRect.Width),
                        (int)Math.Floor(overlayWindow.SelectedRect.Height)
                    );
                    return base.Capture();
                }

                // キャンセル
                return null;
            }

            // キャプチャーアイテムがManualScreenRectCaptureItemでない場合は例外
            throw new InvalidOperationException("CaptureItem is not ManualScreenRectCaptureItem");
        }
    }
}