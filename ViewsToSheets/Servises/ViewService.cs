using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MagicEntry.Plugins.ViewsToSheets.Services
{
    /// <summary>
    /// Сервис для работы с видами в Revit.
    /// Предоставляет методы для получения, фильтрации и проверки видов.
    /// </summary>
    public class ViewService
    {
        #region Public Methods

        /// <summary>
        /// Получает выбранные в диспетчере проекта виды.
        /// </summary>
        /// <param name="uiDoc">UI документ Revit</param>
        /// <returns>Список выбранных видов, которые можно разместить на листе</returns>
        public List<View> GetSelectedViews(UIDocument uiDoc)
        {
            if (uiDoc == null) return new List<View>();

            var selection = uiDoc.Selection;
            var selectedIds = selection.GetElementIds();

            var views = new List<View>();
            foreach (var id in selectedIds)
            {
                var element = uiDoc.Document.GetElement(id);
                if (element is View view && CanBePlacedOnSheet(view))
                {
                    views.Add(view);
                }
            }

            return views;
        }

        /// <summary>
        /// Проверяет, может ли вид быть размещен на листе.
        /// </summary>
        /// <param name="view">Проверяемый вид</param>
        /// <returns>True, если вид можно разместить на листе</returns>
        public bool CanBePlacedOnSheet(View view)
        {
            if (view == null) return false;

            return view.CanBePrinted &&
                   view.ViewType != ViewType.DrawingSheet &&
                   view.ViewType != ViewType.ProjectBrowser &&
                   view.ViewType != ViewType.SystemBrowser &&
                   !IsViewAlreadyOnSheet(view);
        }

        /// <summary>
        /// Получает все виды (viewports) на указанном листе.
        /// </summary>
        /// <param name="doc">Документ Revit</param>
        /// <param name="sheet">Лист для анализа</param>
        /// <returns>Список viewports на листе</returns>
        public List<Viewport> GetViewportsOnSheet(Document doc, ViewSheet sheet)
        {
            if (doc == null || sheet == null) return new List<Viewport>();

            var viewports = new List<Viewport>();
            var viewportIds = sheet.GetAllViewports();

            foreach (var viewportId in viewportIds)
            {
                var viewport = doc.GetElement(viewportId) as Viewport;
                if (viewport != null)
                {
                    viewports.Add(viewport);
                }
            }

            return viewports;
        }

        /// <summary>
        /// Создает строку с именами видов из списка viewports.
        /// </summary>
        /// <param name="doc">Документ Revit</param>
        /// <param name="viewports">Список viewports</param>
        /// <returns>Строка с именами видов, разделенными точками</returns>
        public string GetViewsNamesString(Document doc, List<Viewport> viewports)
        {
            if (doc == null || viewports == null) return string.Empty;

            var names = viewports
                .Select(vp => doc.GetElement(vp.ViewId) as View)
                .Where(v => v != null)
                .Select(v => v.Name)
                .ToList();

            return string.Join(". ", names);
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Проверяет, не размещен ли уже вид на каком-либо листе.
        /// </summary>
        /// <param name="view">Проверяемый вид</param>
        /// <returns>True, если вид уже размещен на листе</returns>
        private bool IsViewAlreadyOnSheet(View view)
        {
            try
            {
                var parameter = view.get_Parameter(BuiltInParameter.VIEWPORT_SHEET_NUMBER);
                return parameter != null && !string.IsNullOrEmpty(parameter.AsString());
            }
            catch
            {
                return false;
            }
        }

        #endregion
    }
}
