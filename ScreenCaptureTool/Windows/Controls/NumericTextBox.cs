using System.Windows.Controls;
using System.Windows.Input;

namespace ScreenCaptureTool.Windows.Controls
{
    /// <summary>
    /// 数値入力のみを受け付けるTextBox
    /// </summary>
    public class NumericTextBox : TextBox
    {
        protected override void OnPreviewTextInput(TextCompositionEventArgs e)
        {
            base.OnPreviewTextInput(e);

            // 数値にパースできない文字列は入力を受け付けない
            e.Handled = !IsTextNumeric(e.Text);
        }

        private bool IsTextNumeric(string text)
        {
            return int.TryParse(text, out _);
        }
    }
}
