using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using MagicEntry.Core.Interfaces;
using MagicEntry.Core.Models;
using MagicEntry.Core.Services;
using MagicEntry.Plugins.ViewsToSheets.Services;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MagicEntry.Plugins.ViewsToSheets
{
    /// <summary>
    /// Команда для создания отдельных листов для каждой основной надписи.
    /// Анализирует активный лист, находит основные надписи и виды внутри них,
    /// создает новые листы и переносит виды с сохранением относительного положения.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CreateSheetsFromTitleBlocksCommand : IExternalCommand, IPlugin
    {
        #region Fields

        private readonly ViewService _viewService;
        private readonly TitleBlockService _titleBlockService;
        private readonly SheetService _sheetService;
        private readonly ViewportPlacementService _placementService;

        #endregion

        #region Constructor

        /// <summary>
        /// Инициализирует новый экземпляр команды создания листов из основных надписей.
        /// </summary>
        public CreateSheetsFromTitleBlocksCommand()
        {
            _viewService = new ViewService();
            _titleBlockService = new TitleBlockService();
            _sheetService = new SheetService();
            _placementService = new ViewportPlacementService();
        }

        #endregion

        #region IPlugin Implementation

        /// <summary>
        /// Конфигурационная информация о плагине, устанавливается системой.
        /// </summary>
        public PluginInfo Info { get; set; }

        /// <summary>
        /// Указывает, активен ли плагин.
        /// </summary>
        public bool IsEnabled { get; set; }

        /// <summary>
        /// Выполняет внутреннюю инициализацию плагина при загрузке системы.
        /// </summary>
        /// <returns>True в случае успешной инициализации</returns>
        public bool Initialize()
        {
            try
            {
                var pathService = ServiceProvider.GetService<IPathService>();
                var initService = ServiceProvider.GetService<IPluginInitializationService>();

                if (pathService == null || initService == null)
                    return false;

                var pluginName = Info?.Name ?? Constants.PLUGIN_NAME;
                return initService.InitializePlugin(pluginName);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Освобождает ресурсы, используемые плагином.
        /// </summary>
        public void Shutdown()
        {
            // Логика завершения работы плагина
        }

        #endregion

        #region IExternalCommand Implementation

        /// <summary>
        /// Точка входа для выполнения команды создания листов из основных надписей.
        /// </summary>
        /// <param name="commandData">Данные команды</param>
        /// <param name="message">Сообщение об ошибке</param>
        /// <param name="elements">Набор элементов</param>
        /// <returns>Результат выполнения команды</returns>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiDoc = commandData.Application.ActiveUIDocument;
                var doc = uiDoc.Document;

                // Валидация входных данных
                var validationResult = ValidateInput(uiDoc, doc, out ViewSheet activeSheet);
                if (validationResult != Result.Succeeded) return validationResult;

                // Получаем данные для обработки
                var processingData = PrepareProcessingData(doc, activeSheet);
                if (processingData == null) return Result.Cancelled;

                // Выполняем создание листов
                return ExecuteSheetCreation(doc, activeSheet, processingData);
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show(Messages.ERROR_TITLE, string.Format(Messages.EXECUTION_ERROR, ex.Message));
                return Result.Failed;
            }
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Валидирует входные данные команды.
        /// </summary>
        /// <param name="uiDoc">UI документ</param>
        /// <param name="doc">Документ</param>
        /// <param name="activeSheet">Активный лист</param>
        /// <returns>Результат валидации</returns>
        private Result ValidateInput(UIDocument uiDoc, Document doc, out ViewSheet activeSheet)
        {
            activeSheet = null;

            // Проверяем, что активный вид - это лист
            if (!(uiDoc.ActiveView is ViewSheet sheet))
            {
                TaskDialog.Show(Messages.ERROR_TITLE, Messages.ACTIVE_VIEW_NOT_SHEET);
                return Result.Cancelled;
            }

            activeSheet = sheet;
            return Result.Succeeded;
        }

        /// <summary>
        /// Подготавливает данные для обработки.
        /// </summary>
        /// <param name="doc">Документ</param>
        /// <param name="activeSheet">Активный лист</param>
        /// <returns>Данные для обработки или null</returns>
        private ProcessingData PrepareProcessingData(Document doc, ViewSheet activeSheet)
        {
            // Получаем основные надписи на листе
            var titleBlocks = _titleBlockService.GetTitleBlocksOnSheet(doc, activeSheet);
            if (titleBlocks == null || !titleBlocks.Any())
            {
                TaskDialog.Show(Messages.ERROR_TITLE, Messages.NO_TITLE_BLOCKS_FOUND);
                return null;
            }

            // Получаем виды на листе
            var viewports = _viewService.GetViewportsOnSheet(doc, activeSheet);
            if (viewports == null || !viewports.Any())
            {
                TaskDialog.Show(Messages.ERROR_TITLE, Messages.NO_VIEWS_ON_SHEET);
                return null;
            }

            // Группируем виды по основным надписям
            var titleBlockGroups = _titleBlockService.GroupViewportsByTitleBlocks(doc, titleBlocks, viewports);
            if (titleBlockGroups.Count <= 1)
            {
                TaskDialog.Show(Messages.ERROR_TITLE, Messages.INSUFFICIENT_TITLE_BLOCKS);
                return null;
            }

            return new ProcessingData
            {
                TitleBlocks = titleBlocks,
                Viewports = viewports,
                TitleBlockGroups = titleBlockGroups
            };
        }

        /// <summary>
        /// Выполняет создание листов.
        /// </summary>
        /// <param name="doc">Документ</param>
        /// <param name="originalSheet">Исходный лист</param>
        /// <param name="data">Данные для обработки</param>
        /// <returns>Результат выполнения</returns>
        private Result ExecuteSheetCreation(Document doc, ViewSheet originalSheet, ProcessingData data)
        {
            using (Transaction trans = new Transaction(doc, Messages.TRANSACTION_NAME_COMMAND3))
            {
                trans.Start();

                int successSheetsCount = CreateSheetsFromTitleBlockGroups(doc, originalSheet, data.TitleBlockGroups);

                if (successSheetsCount > 0)
                {
                    trans.Commit();
                    TaskDialog.Show(Messages.SUCCESS_TITLE,
                        string.Format(Messages.SHEETS_CREATED_SUCCESS, successSheetsCount));
                    return Result.Succeeded;
                }
                else
                {
                    trans.RollBack();
                    TaskDialog.Show(Messages.ERROR_TITLE, Messages.SHEET_CREATION_FAILED);
                    return Result.Failed;
                }
            }
        }

        /// <summary>
        /// Создает новые листы для каждой группы основных надписей и переносит виды.
        /// </summary>
        /// <param name="doc">Документ</param>
        /// <param name="originalSheet">Исходный лист</param>
        /// <param name="titleBlockGroups">Группы основных надписей</param>
        /// <returns>Количество успешно созданных листов</returns>
        private int CreateSheetsFromTitleBlockGroups(Document doc, ViewSheet originalSheet,
            Dictionary<FamilyInstance, List<Viewport>> titleBlockGroups)
        {
            int successSheetsCount = 0;
            try
            {
                var groupsList = titleBlockGroups.ToList();

                // Создаем новые листы для всех групп кроме первой
                for (int i = 1; i < groupsList.Count; i++)
                {
                    if (CreateSingleSheet(doc, originalSheet, titleBlockGroups, i))
                    {
                        successSheetsCount++;
                    }
                }

                // Удаляем лишние основные надписи с исходного листа
                CleanupOriginalSheet(doc, groupsList);

                return successSheetsCount;
            }
            catch
            {
                return successSheetsCount;
            }
        }

        /// <summary>
        /// Создает один лист для указанной группы.
        /// </summary>
        /// <param name="doc">Документ</param>
        /// <param name="originalSheet">Исходный лист</param>
        /// <param name="titleBlockGroups">Все группы</param>
        /// <param name="index">Индекс группы</param>
        /// <returns>True в случае успеха</returns>
        private bool CreateSingleSheet(Document doc, ViewSheet originalSheet,
            Dictionary<FamilyInstance, List<Viewport>> titleBlockGroups, int index)
        {
            var groupsList = titleBlockGroups.ToList();
            FamilyInstance originalTitleBlock = groupsList[index].Key;
            List<Viewport> viewports = groupsList[index].Value;

            if (viewports.Count == 0) return false;

            // Создаем новый лист
            ViewSheet newSheet = _sheetService.CreateNewSheet(doc, originalSheet, originalTitleBlock, viewports);
            if (newSheet == null) return false;

            // Переносим виды на новый лист
            _placementService.MoveViewportsToSheet(doc, titleBlockGroups, index, newSheet);

            // Копируем параметры
            CopyParameters(originalSheet, newSheet, originalTitleBlock);

            return true;
        }

        /// <summary>
        /// Копирует параметры на новый лист.
        /// </summary>
        /// <param name="originalSheet">Исходный лист</param>
        /// <param name="newSheet">Новый лист</param>
        /// <param name="originalTitleBlock">Исходная основная надпись</param>
        private void CopyParameters(ViewSheet originalSheet, ViewSheet newSheet, FamilyInstance originalTitleBlock)
        {
            // Копируем параметры листа
            ParameterCopyService.CopySheetParameters(originalSheet, newSheet);

            // Копируем параметры основной надписи
            var targetTitleBlock = _titleBlockService.GetTitleBlocksOnSheet(newSheet.Document, newSheet).FirstOrDefault();
            if (targetTitleBlock != null)
            {
                ParameterCopyService.CopyTitleBlockParameters(originalTitleBlock, targetTitleBlock);
            }
        }

        /// <summary>
        /// Очищает исходный лист от лишних основных надписей.
        /// </summary>
        /// <param name="doc">Документ</param>
        /// <param name="groupsList">Список групп</param>
        private void CleanupOriginalSheet(Document doc, List<KeyValuePair<FamilyInstance, List<Viewport>>> groupsList)
        {
            for (int i = 1; i < groupsList.Count; i++)
            {
                if (groupsList[i].Value.Count != 0)
                {
                    doc.Delete(groupsList[i].Key.Id);
                }
            }
        }

        #endregion

        #region Helper Classes

        /// <summary>
        /// Класс для хранения данных обработки.
        /// </summary>
        private class ProcessingData
        {
            public List<FamilyInstance> TitleBlocks { get; set; }
            public List<Viewport> Viewports { get; set; }
            public Dictionary<FamilyInstance, List<Viewport>> TitleBlockGroups { get; set; }
        }

        #endregion
    }
}
