using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SaveFileDialog = System.Windows.Forms.SaveFileDialog;

namespace ExportSelectedAs
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiApp = commandData.Application;
            var uiDoc = uiApp.ActiveUIDocument;
            var doc = uiDoc.Document;

            var pickedRef = uiDoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, new ModelGroupSelectionFilter(), "모델그룹 선택");
            if (pickedRef == null)
                return Result.Cancelled;

            Element modelGroup = doc.GetElement(pickedRef.ElementId);

            List<ElementId> elementIds = GetElementsInModelGroup(doc, modelGroup as Group);

            //내보내기
            FolderBrowserDialog fldDlg = new FolderBrowserDialog();
            if (fldDlg.SelectedPath == null)
                return Result.Cancelled;
            else
                ExportToNWC(doc, elementIds, fldDlg.SelectedPath, modelGroup.Name);

            return Result.Succeeded;
        }

        private List<ElementId> GetElementsInModelGroup(Document doc, Group modelGroup)
        {
            List<ElementId> elementIds = new List<ElementId>();

            // 모델 그룹에 포함된 Element들 가져오기
            ICollection<ElementId> modelGroupMemberIds = modelGroup.GetMemberIds();

            foreach (ElementId memberId in modelGroupMemberIds)
            {
                Element member = doc.GetElement(memberId);
                if (member != null)
                    elementIds.Add(memberId);
            }

            return elementIds;
        }

        private string GetElementName(ElementId elemId, Document doc)
        {
            return doc.GetElement(elemId).Name;
        }

        private void ExportToNWC(Document doc, ICollection<ElementId> elementIds, string outputPath, string fileName)
        {
            var option = new NavisworksExportOptions();
            option.SetSelectedElementIds(elementIds);
            try
            {
                doc.Export(outputPath, fileName, option);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message, TaskDialogCommonButtons.Ok);
                return;
            }
        }
    }

    public class ModelGroupSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is Group;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }


}
