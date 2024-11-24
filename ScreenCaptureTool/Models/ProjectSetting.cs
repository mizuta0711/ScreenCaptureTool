using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Xml.Serialization;

namespace ScreenCaptureTool.Models
{
    /// <summary>
    /// プロジェクトの設定を保持するクラス
    /// </summary>
    public class ProjectSetting
    {
        #region Properties

        /// <summary>
        /// プロジェクトファイルのパス
        /// </summary>
        public string FilePath;

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
        /// サムネイル画像サイズ
        /// </summary>
        public int ThumbnailSize { get; set; }

        /// <summary>
        /// 画像保存先フォルダ
        /// </summary>
        public string SaveFolderPath { get; set; }

        /// <summary>
        /// 撮影設定一覧
        /// </summary>
        public ObservableCollection<RecordingSetting> RecordingSettings;

        #endregion Properties

        #region Constructor

        /// <summary>
        /// デフォルトコンストラクタ
        /// </summary>
        public ProjectSetting()
        {
            WindowLeft = 0;
            WindowTop = 0;
            WindowWidth = 0;
            WindowHeight = 0;

            ThumbnailSize = 200;
            SaveFolderPath = Environment.CurrentDirectory;
            RecordingSettings = new ObservableCollection<RecordingSetting>();

            FilePath = Path.Combine(Environment.CurrentDirectory, "ScreenCaptureTool.scp");
        }

        #endregion Constructor

        #region Methods

        #region Methods(Public)

        /// <summary>
        /// ファイルに保存する
        /// </summary>
        /// <param name="filePath">ファイルパス</param>
        /// <returns>true: 成功 / false: 失敗</returns>
        public bool Save(string filePath)
        {
            try
            {
                var serializer = new XmlSerializer(typeof(ProjectSetting));
                using (var fs = new FileStream(filePath, FileMode.Create))
                {
                    serializer.Serialize(fs, this);
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        #endregion Methods(Public)

        #region Methods(Static)

        /// <summary>
        /// ファイルから設定を読み込む
        /// </summary>
        /// <param name="filePath">ファイルパス</param>
        /// <returns>インスタンス(失敗時はnull)</returns>
        public static ProjectSetting? Load(string filePath)
        {
            try
            {
                var serializer = new XmlSerializer(typeof(ProjectSetting));
                using (var fs = new FileStream(filePath, FileMode.Open))
                {
                    var projectSetting = serializer.Deserialize(fs) as ProjectSetting;
                    if (projectSetting != null)
                    {
                        projectSetting.FilePath = filePath;
                        return projectSetting;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return null;
        }

        #endregion Methods(Static)

        #endregion Methods
    }
}