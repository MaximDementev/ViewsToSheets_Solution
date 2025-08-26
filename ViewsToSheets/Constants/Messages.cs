namespace MagicEntry.Plugins.ViewsToSheets
{
    /// <summary>
    /// Статический класс для хранения всех текстовых сообщений плагина.
    /// Централизует управление текстом интерфейса и сообщений об ошибках.
    /// </summary>
    public static class Messages
    {
        #region Dialog Titles
        public const string ERROR_TITLE = "Ошибка";
        public const string INFO_TITLE = "Информация";
        public const string WARNING_TITLE = "Предупреждение";
        public const string SUCCESS_TITLE = "Успешно";
        #endregion

        #region Main Window
        public const string MAIN_WINDOW_TITLE = "Размещение видов на листах";
        public const string COMMAND_1_TITLE = "Команда 1: Размещение в столбик";
        public const string COMMAND_1_DESCRIPTION = "Размещение выбранных видов в столбик на одном листе";
        public const string COMMAND_2_TITLE = "Команда 2: Автоматическое распределение";
        public const string COMMAND_2_DESCRIPTION = "Автоматическое распределение видов по основным надписям";
        public const string COMMAND_3_TITLE = "Команда 3: Создание отдельных листов";
        public const string COMMAND_3_DESCRIPTION = "Создание отдельного листа для каждой основной надписи";
        #endregion

        #region Error Messages
        public const string NO_ACTIVE_DOCUMENT = "Нет активного документа Revit";
        public const string EXECUTION_ERROR = "Ошибка при выполнении плагина: {0}";
        public const string INVALID_COMMAND = "Неверная команда";
        public const string COMMAND_NOT_IMPLEMENTED = "Команда еще не реализована";
        public const string NO_VIEWS_SELECTED = "Не выбраны виды для размещения";
        public const string NO_SHEET_SELECTED = "Не выбран лист для размещения";
        public const string INVALID_SPACING = "Неверное значение расстояния между видами";
        public const string ACTIVE_VIEW_NOT_SHEET = "Активный вид должен быть листом";
        public const string NO_TITLE_BLOCKS_FOUND = "На листе не найдены основные надписи";
        public const string NO_VIEWS_ON_SHEET = "На листе нет размещенных видов";
        public const string INSUFFICIENT_TITLE_BLOCKS = "Недостаточно основных надписей для создания отдельных листов";
        public const string SHEET_CREATION_FAILED = "Не удалось создать новые листы";
        #endregion

        #region Success Messages
        public const string VIEWS_PLACED_SUCCESSFULLY = "Виды успешно размещены на листе";
        public const string SHEETS_CREATED_SUCCESS = "Успешно создано {0} новых листов";
        #endregion

        #region Command 1 Messages
        public const string SELECT_VIEWS_PROMPT = "Выберите виды для размещения";
        public const string SELECT_TARGET_SHEET = "Выберите лист для размещения видов";
        public const string SPACING_DIALOG_TITLE = "Расстояние между видами";
        public const string SPACING_INSTRUCTION = "Выберите расстояние между видами:";
        public const string SPACING_DIALOG_CONTENT = "Выберите один из предустановленных вариантов расстояния между видами на листе.";
        public const string SPACING_OPTION_SMALL = "Малое расстояние ({0} мм)";
        public const string SPACING_OPTION_MEDIUM = "Среднее расстояние ({0} мм)";
        public const string SPACING_OPTION_LARGE = "Большое расстояние ({0} мм)";
        public const string VIEWS_PLACED_SUCCESS = "Успешно размещено {0} видов на листе '{1}'";
        public const string PLACEMENT_FAILED = "Не удалось разместить виды на листе";
        public const string TRANSACTION_NAME = "Размещение видов в столбик";
        #endregion

        #region Command 3 Messages
        public const string TRANSACTION_NAME_COMMAND3 = "Создание листов из основных надписей";
        #endregion
    }
}
