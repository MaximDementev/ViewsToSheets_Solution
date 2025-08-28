using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MagicEntry.Plugins.ViewsToSheets.Services
{
    /// <summary>
    /// Сервис для работы с листами в Revit.
    /// Предоставляет методы для создания листов и работы с их свойствами.
    /// </summary>
    public class SheetService
    {
        #region Public Methods

        /// <summary>
        /// Создает новый лист на основе исходного.
        /// </summary>
        /// <param name="doc">Документ Revit</param>
        /// <param name="originalSheet">Исходный лист</param>
        /// <param name="titleBlock">Основная надпись для нового листа</param>
        /// <param name="viewports">Список viewports для определения имени</param>
        /// <param name="index">Индекс для создания уникального номера</param>
        /// <returns>Новый лист или null в случае ошибки</returns>
        public ViewSheet CreateNewSheet(Document doc, ViewSheet originalSheet, FamilyInstance titleBlock,
            List<Viewport> viewports)
        {
            if (doc == null || originalSheet == null || titleBlock == null) return null;

            try
            {
                var sheetTypeId = originalSheet.GetTypeId();
                var sheetType = doc.GetElement(sheetTypeId) as ElementType;
                if (sheetType == null) return null;

                // Получаем ID типа основной надписи
                ElementId titleBlockTypeId = GetTitleBlockTypeId(doc, titleBlock);
                if (titleBlockTypeId == ElementId.InvalidElementId) return null;

                // Создаем лист
                ViewSheet newSheet = ViewSheet.Create(doc, titleBlockTypeId);
                if (newSheet == null) return null;

                // Устанавливаем уникальные номер и имя
                var viewService = new ViewService();
                string sheetName = viewService.GetViewsNamesString(doc, viewports);
                sheetName = GetUniqueViewName(doc, sheetName);
                string sheetNumber = $"{originalSheet.SheetNumber}";
                sheetNumber = GetUniqueSheetNumber(doc, sheetNumber);

                newSheet.SheetNumber = sheetNumber;
                newSheet.Name = sheetName;

                return newSheet;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка создания листа: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Получает уникальное имя для вида.
        /// </summary>
        /// <param name="doc">Документ Revit</param>
        /// <param name="baseName">Базовое имя</param>
        /// <returns>Уникальное имя</returns>
        public string GetUniqueViewName(Document doc, string baseName)
        {
            if (doc == null || string.IsNullOrEmpty(baseName)) return "Новый лист";

            var existingNames = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Select(v => v.Name)
                .ToHashSet();

            if (!existingNames.Contains(baseName))
                return baseName;

            int counter = 1;
            string newName;
            do
            {
                newName = $"{baseName} ({counter})";
                counter++;
            }
            while (existingNames.Contains(newName));

            return newName;
        }

        /// <summary>
        /// Получает уникальный номер для листа.
        /// </summary>
        /// <param name="doc">Документ Revit</param>
        /// <param name="baseNumber">Базовый номер</param>
        /// <returns>Уникальный номер</returns>
        public string GetUniqueSheetNumber(Document doc, string baseNumber)
        {
            if (doc == null || string.IsNullOrEmpty(baseNumber)) return "001";

            var existingNumbers = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Select(s => s.SheetNumber)
                .ToHashSet();

            if (!existingNumbers.Contains(baseNumber))
                return baseNumber;

            int counter = 1;
            string newNumber;
            do
            {
                newNumber = $"{baseNumber} -({counter})";
                counter++;
            }
            while (existingNumbers.Contains(newNumber));

            return newNumber;
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Получает ID типа основной надписи.
        /// </summary>
        /// <param name="doc">Документ Revit</param>
        /// <param name="titleBlock">Основная надпись</param>
        /// <returns>ID типа основной надписи</returns>
        private ElementId GetTitleBlockTypeId(Document doc, FamilyInstance titleBlock)
        {
            ElementId titleBlockTypeId = titleBlock.GetTypeId();
            if (titleBlockTypeId != ElementId.InvalidElementId)
                return titleBlockTypeId;

            // Берем случайный тип основной надписи
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            return collector
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .OfClass(typeof(FamilySymbol))
                .FirstElementId();
        }

        #endregion
    }
}
