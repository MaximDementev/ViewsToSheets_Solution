namespace MagicEntry.Plugins.ViewsToSheets
{
    /// <summary>
    /// Статический класс для хранения постоянных данных плагина ViewsToSheets.
    /// Содержит названия, настройки по умолчанию и другие константы.
    /// </summary>
    public static class Constants
    {
        #region Plugin Info
        public const string PLUGIN_NAME = "ViewsToSheets";
        public const string PLUGIN_DISPLAY_NAME = "Размещение видов на листах";
        public const string PLUGIN_VERSION = "1.0.0";
        #endregion

        #region Default Settings
        public const double SPACING_SMALL = 5.0; // мм
        public const double SPACING_MEDIUM = 10.0; // мм  
        public const double SPACING_LARGE = 20.0; // мм
        public const int DEFAULT_MAX_VIEWS_PER_SHEET = 10;
        public const double DEFAULT_MIN_SHEET_SIZE = 210.0; // A4 width in mm
        public const double DEFAULT_MAX_SHEET_SIZE = 841.0; // A0 width in mm
        #endregion

        #region Start Position
        public const double START_POSITION_X = 0.5; // футы
        public const double START_POSITION_Y = 0.5; // футы
        #endregion

        #region Sheet Formats
        public const string FORMAT_A4 = "A4";
        public const string FORMAT_A3 = "A3";
        public const string FORMAT_A2 = "A2";
        public const string FORMAT_A1 = "A1";
        public const string FORMAT_A0 = "A0";
        #endregion

        #region Units
        public const double MM_TO_FEET = 0.00328084; // Conversion factor from mm to feet
        public const double FEET_TO_MM = 304.8; // Conversion factor from feet to mm
        #endregion
    }
}
