using System;
using System.Text.RegularExpressions;

namespace ScreenCaptureTool.Utilities
{
    /// <summary>
    /// ファイル名フォーマット
    /// </summary>
    public class FileNameFormatter
    {
        #region Fields

        /// <summary>
        /// 名称
        /// </summary>
        private string _name;

        /// <summary>
        /// 拡張子
        /// </summary>
        private string _extension;

        #endregion Fields

        #region Properties

        /// <summary>
        /// 書式文字列
        /// </summary>
        public string Format { get; private set; }

        /// <summary>
        /// フォーマットされた名称
        /// </summary>
        public string FormattedName
        {
            get
            {
                var filename = Format + _extension;
                // %NAME%：名称に置換
                filename = filename.Replace("%NAME%", _name);
                // %DATETIME%：日時に置換
                filename = filename.Replace("%DATETIME%", DateTime.Now.ToString("yyyyMMdd_HHmmss"));
                // %DATE%：日付に置換
                filename = filename.Replace("%DATE%", DateTime.Now.ToString("yyyyMMdd"));
                // %TIME%：時刻に置換
                filename = filename.Replace("%TIME%", DateTime.Now.ToString("HHmmss"));
                // %NUMBERS{開始}.桁数%：連番に置換
                var match = Regex.Match(filename, "%NUMBER{([0-9]+)}.([0-9]+)%");
                if (match.Success && match.Groups.Count == 3)
                {
                    // 連番を生成
                    var start = int.Parse(match.Groups[1].Value);
                    var digits = int.Parse(match.Groups[2].Value);
                    var number = start;
                    filename = filename.Replace(match.Value, number.ToString().PadLeft(digits, '0'));
                }
                return filename;
            }
        }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="format">書式文字列</param>
        /// <param name="name">名称</param>
        /// <param name="extension">拡張子</param>
        public FileNameFormatter(string format, string name, string extension)
        {
            Format = format;
            _name = name;
            _extension = extension;
        }

        #endregion Constructors

        #region Methods

        /// <summary>
        /// ファイル名フォーマットの連番をインクリメントする
        /// </summary>
        /// <returns>インクリメント後のフォーマット</returns>
        public string IncriementNumber()
        {
            // 「%NUMBERS{開始}.桁数%」の指定がある場合はインクリメント
            var match = Regex.Match(Format, "%NUMBER{([0-9]+)}.([0-9]+)%");
            if (match.Success && match.Groups.Count == 3)
            {
                // 連番をインクリメントした書式に更新
                var start = int.Parse(match.Groups[1].Value);
                var digits = int.Parse(match.Groups[2].Value);
                var number = start + 1;
                Format = Format.Replace(match.Value, $"%NUMBER{{{number}}}.{digits}%");
            }
            return Format;
        }

        #endregion Methods
    }
}