using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ViewsToSheets.Servises
{
    public static class ParameterCopyService
    {
        /// <summary>
        /// Копирует параметры экземпляра с исходного листа на целевой.
        /// </summary>
        public static void CopySheetParameters(ViewSheet source, ViewSheet target)
        {
            if (source == null || target == null) return;
            CopyParameters(source, target);
        }

        /// <summary>
        /// Копирует параметры экземпляра с исходной основной надписи на целевую.
        /// </summary>
        public static void CopyTitleBlockParameters(FamilyInstance source, FamilyInstance target)
        {
            if (source == null || target == null) return;
            CopyParameters(source, target);
        }

        private static void CopyParameters(Element source, Element target)
        {
            foreach (Parameter srcParam in source.Parameters)
            {
                if (srcParam.IsReadOnly) continue;
                if (srcParam.StorageType == StorageType.None) continue;
                if (srcParam.Definition == null) continue;

                string paramName = srcParam.Definition.Name;

                // для листов пропускаем Имя и Номер
                if (
                    paramName.Equals("Номер листа", StringComparison.OrdinalIgnoreCase) ||
                     paramName.Equals("Sheet Number", StringComparison.OrdinalIgnoreCase) ||
                     paramName.Equals("Имя", StringComparison.OrdinalIgnoreCase) ||
                     paramName.Equals("Sheet Name", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                Parameter trgParam = target.LookupParameter(paramName);
                if (trgParam == null || trgParam.IsReadOnly) continue;

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
        }

    }
}
