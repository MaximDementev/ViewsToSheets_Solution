using Autodesk.Revit.DB;
using System;

namespace MagicEntry.Plugins.ViewsToSheets.Services
{
    /// <summary>
    /// Сервис для копирования параметров между элементами Revit.
    /// Предоставляет методы для копирования параметров листов и основных надписей.
    /// </summary>
    public static class ParameterCopyService
    {
        #region Public Methods

        /// <summary>
        /// Копирует параметры экземпляра с исходного листа на целевой.
        /// </summary>
        /// <param name="source">Исходный лист</param>
        /// <param name="target">Целевой лист</param>
        public static void CopySheetParameters(ViewSheet source, ViewSheet target)
        {
            if (source == null || target == null) return;
            CopyParameters(source, target);
        }

        /// <summary>
        /// Копирует параметры экземпляра с исходной основной надписи на целевую.
        /// </summary>
        /// <param name="source">Исходная основная надпись</param>
        /// <param name="target">Целевая основная надпись</param>
        public static void CopyTitleBlockParameters(FamilyInstance source, FamilyInstance target)
        {
            if (source == null || target == null) return;
            CopyParameters(source, target);
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Копирует параметры между элементами.
        /// </summary>
        /// <param name="source">Исходный элемент</param>
        /// <param name="target">Целевой элемент</param>
        private static void CopyParameters(Element source, Element target)
        {
            foreach (Parameter srcParam in source.Parameters)
            {
                if (!IsParameterCopyable(srcParam)) continue;

                string paramName = srcParam.Definition.Name;

                // Пропускаем системные параметры листов
                if (IsSystemSheetParameter(paramName)) continue;

                Parameter trgParam = target.LookupParameter(paramName);
                if (trgParam == null || trgParam.IsReadOnly) continue;

                CopyParameterValue(srcParam, trgParam, paramName);
            }
        }

        /// <summary>
        /// Проверяет, можно ли копировать параметр.
        /// </summary>
        /// <param name="parameter">Проверяемый параметр</param>
        /// <returns>True, если параметр можно копировать</returns>
        private static bool IsParameterCopyable(Parameter parameter)
        {
            return !parameter.IsReadOnly &&
                   parameter.StorageType != StorageType.None &&
                   parameter.Definition != null;
        }

        /// <summary>
        /// Проверяет, является ли параметр системным параметром листа.
        /// </summary>
        /// <param name="paramName">Имя параметра</param>
        /// <returns>True, если параметр системный</returns>
        private static bool IsSystemSheetParameter(string paramName)
        {
            return paramName.Equals("Номер листа", StringComparison.OrdinalIgnoreCase) ||
                   paramName.Equals("Sheet Number", StringComparison.OrdinalIgnoreCase) ||
                   paramName.Equals("Имя", StringComparison.OrdinalIgnoreCase) ||
                   paramName.Equals("Sheet Name", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Копирует значение параметра.
        /// </summary>
        /// <param name="srcParam">Исходный параметр</param>
        /// <param name="trgParam">Целевой параметр</param>
        /// <param name="paramName">Имя параметра для логирования</param>
        private static void CopyParameterValue(Parameter srcParam, Parameter trgParam, string paramName)
        {
            try
            {
                switch (srcParam.StorageType)
                {
                    case StorageType.String:
                        trgParam.Set(srcParam.AsString());
                        break;
                    case StorageType.Double:
                        trgParam.Set(srcParam.AsDouble());
                        break;
                    case StorageType.Integer:
                        trgParam.Set(srcParam.AsInteger());
                        break;
                    case StorageType.ElementId:
                        trgParam.Set(srcParam.AsElementId());
                        break;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"Не удалось скопировать параметр {paramName}: {ex.Message}");
            }
        }

        #endregion
    }
}
