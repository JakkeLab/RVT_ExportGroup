using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
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

            //선택모드
            IList<Reference> pickedRefs;
            try
            {
                pickedRefs = uiDoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, new ModelGroupSelectionFilter(), "모델그룹 선택");
            }
            catch (Exception ex)
            {
                return Result.Cancelled;
            }

            if (pickedRefs == null)
                return Result.Cancelled;

            //내보내기
            FolderBrowserDialog fldDlg = new FolderBrowserDialog();
            DialogResult result = fldDlg.ShowDialog();
            if (result != DialogResult.OK || string.IsNullOrEmpty(fldDlg.SelectedPath))
                return Result.Cancelled;

            //파일 있는지 한번 검사
            string folderPath = fldDlg.SelectedPath;
            var fileNames = Directory.GetFiles(folderPath).ToList();
            List<string> fileNamesToExport = new List<string>();
            foreach (Reference pickedRef in pickedRefs)
            {
                Element modelGroup = doc.GetElement(pickedRef.ElementId);
                List<ElementId> elementIds = GetElementsInModelGroup(doc, modelGroup as Group);
                fileNamesToExport.Add(Path.Combine(modelGroup.Name + ".nwc"));
            }

            //이미 있는경우 
            if(fileNames.Count(x => fileNamesToExport.Contains(x)) != 0)
            {
                DialogResult msgResult = MessageBox.Show("일부 파일이 이미 폴더에 있습니다. 계속하시겠습니까?", "파일 덮어쓰기", MessageBoxButtons.YesNo);
                if(msgResult == DialogResult.No)
                {
                    return Result.Cancelled;
                }
            }

            foreach (Reference pickedRef in pickedRefs)
            {
                Element modelGroup = doc.GetElement(pickedRef.ElementId);
                List<ElementId> elementIds = GetElementsInModelGroup(doc, modelGroup as Group);
                string filePath = Path.Combine(fldDlg.SelectedPath, modelGroup.Name + ".nwc");
                if(File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
                ExportToNWC(doc, elementIds, fldDlg.SelectedPath, modelGroup.Name);
            }
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
            option.ExportScope = NavisworksExportScope.SelectedElements;
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
