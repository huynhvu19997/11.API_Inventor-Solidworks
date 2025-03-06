using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.IO;

namespace OOP_InAPI
{
    class Program
    {
        static void Main(string[] args)
        {
            inventorAPI inventorAPI = new inventorAPI();
            inventorAPI.Run();
        }
    }

    public class inventorAPI
    {
        private Inventor.Application invApp;
        private static bool isPicking = false;

        [DllImport("user32.dll")]
        public static extern short GetAsyncKeyState(int vKey);

        [STAThread]

        public void Run()
        {
            if (!CheckInventorRunning())
            {
                DialogResult result = MessageBox.Show("Bạn có muốn mở phần mềm không?", "Inventor", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    startInventor();
                }
                else
                {
                    Console.WriteLine("Phần mềm đã bị đóng.");
                    return;
                }
            }

            //CreateNewFile();
            //CreatePart(10,20,30,10);
            //DrawLine(10,20);
            //CreateRectangularBlockWithCircularCut(100,50,30,20);
            //ChangeParameterInPart("p", "40");
            //CreateInventorProject(@"D:\1. OE_Keson\6.2025\4.IVentor\test", "ABC1Project");
            //PitchEdgeIfPartOpen1();
            //SelectObjectsAndCreateProject();
        }


        private bool CheckInventorRunning()
        {
            try
            {
                invApp = Marshal.GetActiveObject("Inventor.Application") as Inventor.Application; // cách ép kiểu kiểm tra as
                if (invApp != null)
                {
                    Console.WriteLine("phần mềm đã chạy");
                    return true;
                }
            }
            catch (COMException)
            {
                Console.WriteLine(" phần mềm chưa chạy");
            }
            return false;
        }

        private void startInventor()
        {
            try
            {
                Type inventorAppType = Type.GetTypeFromProgID("Inventor.Application");
                invApp = (Inventor.Application)Activator.CreateInstance(inventorAppType);
                invApp.Visible = true;
                Console.WriteLine("phần mềm đã mở");
            }
            catch (Exception ex)
            {
                MessageBox.Show("lỗi khởi động phần mềm" + ex.Message);
            }
        }

        private void stopInventor()
        {
            try
            {
                invApp.Quit();
                Marshal.FinalReleaseComObject(invApp);
                invApp = null;
                Console.WriteLine("đã đóng phần mềm thành công");
            }
            catch (Exception ex)
            {
                MessageBox.Show("lỗi tắt phần mềm" + ex.Message);
            }


        }

        private void CreateNewFile()
        {
            string partTemplate = invApp.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject, SystemOfMeasureEnum.kMetricSystemOfMeasure);
            string assemblyTemplate = invApp.FileManager.GetTemplateFile(DocumentTypeEnum.kAssemblyDocumentObject, SystemOfMeasureEnum.kMetricSystemOfMeasure);
            string drawingTemplate = invApp.FileManager.GetTemplateFile(DocumentTypeEnum.kDrawingDocumentObject, SystemOfMeasureEnum.kMetricSystemOfMeasure);

            //tạo một file part
            PartDocument partDoc = (PartDocument)invApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, partTemplate); // kiểm tra lại có phải là hệ metric với kích thước mm
            partDoc.UnitsOfMeasure.LengthUnits = UnitsTypeEnum.kMillimeterLengthUnits;// chuyển sang mm
            Console.WriteLine("Đã tạo file Part.");
            //tạo một file Assembly
            AssemblyDocument assemDoc = (AssemblyDocument)invApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, assemblyTemplate);
            assemDoc.UnitsOfMeasure.LengthUnits = UnitsTypeEnum.kMillimeterLengthUnits;
            Console.WriteLine("đã tạo file Assembly");
            //Tạo một file drawing
            DrawingDocument drawingDoc = (DrawingDocument)invApp.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, drawingTemplate);
            drawingDoc.UnitsOfMeasure.LengthUnits = UnitsTypeEnum.kMillimeterLengthUnits;
            //AddTitleBlockAndBorder(drawingDoc); // cần có tính kế thừa :InventorDocument
            Console.WriteLine("đã tạo file drawing");


        }

        private void CheckOpenDocuments()
        {
            if (invApp.Documents.Count == 0)
            {
                MessageBox.Show("Chưa mở bất cứ file nào.");
                return;
            }

            // lấy danh sách các tài liệu đang mở
            var openParts = invApp.Documents.OfType<PartDocument>().Select(d => d.DisplayName).ToList();// ghi rõ ý nghĩa của nó
            var openAssemblies = invApp.Documents.OfType<AssemblyDocument>().Select(d => d.DisplayName).ToList();// ghi rõ ý nghĩa của nó
            var openDrawings = invApp.Documents.OfType<DrawingDocument>().Select(d => d.DisplayName).ToList();// ghi rõ ý nghĩa của nó

            //in ra danh sách các tài liệu đang mở
            Console.WriteLine("Các tài liệu đang mở");

            //kiểm tra các file part đang mở và in ra list
            if (openParts.Count > 0)
            {
                Console.WriteLine("Part Documents:");
                openParts.ForEach(name => Console.WriteLine(" - " + name));
            }
            //kiểm tra các file assembly đang mở và in ra list
            if (openAssemblies.Count > 0)
            {
                Console.WriteLine("Assembly Document:");
                openAssemblies.ForEach(name => Console.WriteLine(" - " + name));
            }

            //kiểm tra các file drawing đang mở và in ra list
            if (openDrawings.Count > 0)
            {
                Console.WriteLine("Drawing Document:");
                openDrawings.ForEach(name => Console.WriteLine(" - " + name));
            }

            // Kiểm tra trùng tên tài liệu
            string input = "Nhập tên để kiểm tra trùng tên:" + "Kiểm tra trùng tên" + ""; // có thể nhập từ textbox_combobox
            if (!string.IsNullOrEmpty(input))
            {
                bool exists = openParts.Contains(input) || openAssemblies.Contains(input) || openDrawings.Contains(input);
                if (exists)
                {
                    MessageBox.Show("Tên tài liệu đã tồn tại.");
                }
                else
                {
                    MessageBox.Show("Tên tài liệu không tồn tại.");
                }
            }
        }

        private void FindDocument(string name)
        {
            Document foundDocument = null;

            // Kiểm tra tài liệu Part
            var partDocument = invApp.Documents.OfType<PartDocument>().FirstOrDefault(d => d.DisplayName == name);
            if (partDocument != null)
            {
                foundDocument = (Document)partDocument;
            }
            else
            {
                // Kiểm tra tài liệu Assembly
                var assemblyDocument = invApp.Documents.OfType<AssemblyDocument>().FirstOrDefault(d => d.DisplayName == name);
                if (assemblyDocument != null)
                {
                    foundDocument = (Document)assemblyDocument;
                }
                else
                {
                    // Kiểm tra tài liệu Drawing
                    var drawingDocument = invApp.Documents.OfType<DrawingDocument>().FirstOrDefault(d => d.DisplayName == name);
                    if (drawingDocument != null)
                    {
                        foundDocument = (Document)drawingDocument;
                    }
                }
            }

            if (foundDocument != null)
            {
                foundDocument.Activate();
                MessageBox.Show($"Đã mở tài liệu: {name}");
            }
            else
            {
                MessageBox.Show($"Không tìm thấy tài liệu: {name}");
            }
        }

        private void FindAndOpenDocument(string fileName)
        {
            /// Thư mục dự án
            string projectPath = @"C:\path\to\your\project";

            // Tìm và mở file trong thư mục dự án
            string[] fileExtensions = { "*.ipt", "*.iam", "*.idw" };
            foreach (var extension in fileExtensions)
            {
                string[] files = Directory.GetFiles(projectPath, extension, SearchOption.AllDirectories);
                foreach (var file in files)
                {
                    if (System.IO.Path.GetFileName(file).Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        //OpenDocument(file);
                        try
                        {
                            invApp.Documents.Open(file);
                            MessageBox.Show($"Đã mở tài liệu: {file}");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Lỗi khi mở tài liệu: {ex.Message}");
                        }
                        return;
                    }
                }
            }

            MessageBox.Show($"Không tìm thấy tài liệu: {fileName}");
        }

        private void OpenDocument(string filePath)
        {
            try
            {
                invApp.Documents.Open(filePath);
                MessageBox.Show($"Đã mở tài liệu: {filePath}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi mở tài liệu: {ex.Message}");
            }

        }

        private void DrawLine(double length, double width)
        {
            try
            {


                //// Tạo tài liệu phần
                //PartDocument partDoc = (PartDocument)invApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject,
                //    invApp.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject));
                //PartComponentDefinition partCompDef = partDoc.ComponentDefinition;
                //TransientGeometry transGeom = invApp.TransientGeometry;

                //// Tạo phác thảo
                //PlanarSketch sketch = partCompDef.Sketches.Add(partCompDef.WorkPlanes[3]);
                //SketchLine line = sketch.SketchLines.AddByTwoPoints(transGeom.CreatePoint2d(0, 0), transGeom.CreatePoint2d(length, 0));

                //// Thêm ràng buộc kích thước cho đường thẳng
                //TwoPointDistanceDimConstraint dimConstraint = sketch.DimensionConstraints.AddTwoPointDistance(
                //    line.StartSketchPoint, line.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, transGeom.CreatePoint2d(length / 2, -10), true);
                //dimConstraint.Parameter._Value = length;

                // Tạo tài liệu phần
                PartDocument partDoc = (PartDocument)invApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, invApp.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject));
                partDoc.UnitsOfMeasure.LengthUnits = UnitsTypeEnum.kMillimeterLengthUnits;// chuyển sang mm
                PartComponentDefinition partCompDef = partDoc.ComponentDefinition;
                TransientGeometry transGeom = invApp.TransientGeometry;

                // Tạo phác thảo
                PlanarSketch sketch = partCompDef.Sketches.Add(partCompDef.WorkPlanes[3]);

                // Vẽ các đường thẳng tạo nên hình chữ nhật
                SketchLine line1 = sketch.SketchLines.AddByTwoPoints(transGeom.CreatePoint2d(0, 0), transGeom.CreatePoint2d(length, 0));
                SketchLine line2 = sketch.SketchLines.AddByTwoPoints(line1.EndSketchPoint, transGeom.CreatePoint2d(length, width));
                SketchLine line3 = sketch.SketchLines.AddByTwoPoints(line2.EndSketchPoint, transGeom.CreatePoint2d(0, width));
                SketchLine line4 = sketch.SketchLines.AddByTwoPoints(line3.EndSketchPoint, line1.StartSketchPoint);

                // Thêm ràng buộc vuông góc và song song
                sketch.GeometricConstraints.AddParallel((SketchEntity)line1, (SketchEntity)line3);
                sketch.GeometricConstraints.AddParallel((SketchEntity)line2, (SketchEntity)line4);
                sketch.GeometricConstraints.AddPerpendicular((SketchEntity)line1, (SketchEntity)line2);

                // Thêm ràng buộc kích thước cho các cạnh hình chữ nhật
                TwoPointDistanceDimConstraint lengthDimConstraint = sketch.DimensionConstraints.AddTwoPointDistance(
                    line1.StartSketchPoint, line1.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, transGeom.CreatePoint2d(length / 2, -10), false);

                TwoPointDistanceDimConstraint widthDimConstraint = sketch.DimensionConstraints.AddTwoPointDistance(
                    line2.StartSketchPoint, line2.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, transGeom.CreatePoint2d(length + 10, width / 2), false);

                // Đổi tên tham số mô hình và đặt giá trị cho chúng
                if (lengthDimConstraint.Parameter is ModelParameter lengthParam)
                {
                    lengthParam.Name = "Length";
                    lengthParam.Expression = length + " mm";
                }

                if (widthDimConstraint.Parameter is ModelParameter widthParam)
                {
                    widthParam.Name = "Width";
                    widthParam.Expression = width + " mm";
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("Error inside DrawLineWithModelParameter: " + ex.Message);

            }
        }

        public void CreateRectangularBlockWithCircularCut(double length, double width, double height, double holeDiameter)
        {
           
                // Tạo tài liệu phần
                PartDocument partDoc = (PartDocument)invApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, invApp.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject));
                partDoc.UnitsOfMeasure.LengthUnits = UnitsTypeEnum.kMillimeterLengthUnits; // Chuyển sang mm
                PartComponentDefinition partCompDef = partDoc.ComponentDefinition;
                TransientGeometry transGeom = invApp.TransientGeometry;

                // Tạo phác thảo cho khối hình chữ nhật
                PlanarSketch basesketch = partCompDef.Sketches.Add(partCompDef.WorkPlanes[3]);

                // Vẽ các đường thẳng tạo nên hình chữ nhật
                SketchLine line1 = basesketch.SketchLines.AddByTwoPoints(transGeom.CreatePoint2d(0, 0), transGeom.CreatePoint2d(length, 0));
                SketchLine line2 = basesketch.SketchLines.AddByTwoPoints(line1.EndSketchPoint, transGeom.CreatePoint2d(length, width));
                SketchLine line3 = basesketch.SketchLines.AddByTwoPoints(line2.EndSketchPoint, transGeom.CreatePoint2d(0, width));
                SketchLine line4 = basesketch.SketchLines.AddByTwoPoints(line3.EndSketchPoint, line1.StartSketchPoint);

                // Thêm ràng buộc vuông góc và song song cho các cạnh của khối hình chữ nhật
                basesketch.GeometricConstraints.AddPerpendicular((SketchEntity)line1, (SketchEntity)line2);
                basesketch.GeometricConstraints.AddPerpendicular((SketchEntity)line2, (SketchEntity)line3);
                basesketch.GeometricConstraints.AddPerpendicular((SketchEntity)line3,(SketchEntity)line4);
                basesketch.GeometricConstraints.AddPerpendicular((SketchEntity)line4, (SketchEntity)line1);

                // Thêm ràng buộc kích thước cho các cạnh của khối hình chữ nhật
                TwoPointDistanceDimConstraint lengthDimConstraint = basesketch.DimensionConstraints.AddTwoPointDistance(
                    line1.StartSketchPoint, line1.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, transGeom.CreatePoint2d(length / 2, -10), false);
                //lengthDimConstraint.Parameter._Value = length;

                TwoPointDistanceDimConstraint widthDimConstraint = basesketch.DimensionConstraints.AddTwoPointDistance(
                    line2.StartSketchPoint, line2.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, transGeom.CreatePoint2d(length + 10, width / 2), false);
                //widthDimConstraint.Parameter._Value = width;

                // Tạo khối hình hộp chữ nhật
                Profile baseProfile = basesketch.Profiles.AddForSolid();
                ExtrudeFeature blockExtrude = partCompDef.Features.ExtrudeFeatures.AddByDistanceExtent(
                    baseProfile, height, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection, PartFeatureOperationEnum.kNewBodyOperation);

                // Tạo phác thảo trên bề mặt của khối hình chữ nhật để cắt lỗ tròn
                PlanarSketch cutSketch = partCompDef.Sketches.Add(blockExtrude.Faces[1]);
                SketchCircle cutCircle = cutSketch.SketchCircles.AddByCenterRadius(transGeom.CreatePoint2d(length / 2, width / 2), holeDiameter / 2);

                // Tạo ràng buộc kích thước cho lỗ tròn
                cutSketch.DimensionConstraints.AddDiameter((SketchEntity)cutCircle, transGeom.CreatePoint2d(length / 2, (width / 2) + holeDiameter / 2), false);

                // Cắt lỗ tròn với chiều sâu bằng 1/2 chiều cao của khối hình hộp chữ nhật
                Profile cutProfile = cutSketch.Profiles.AddForSolid();
                ExtrudeFeature cutExtrude = partCompDef.Features.ExtrudeFeatures.AddByDistanceExtent(
                    cutProfile, height / 2, PartFeatureExtentDirectionEnum.kNegativeExtentDirection, PartFeatureOperationEnum.kCutOperation);

            


        }
        /// <summary>
        //Dành cho part
        /// </summary>
        /// <param name="parameterName"></param>
        /// <param name="newValue"></param>

        private void ChangeParameterInPart(string parameterName, string newValue)
        {
            try
            {
                PartDocument oPartDoc;
                // kiểm tra phần mềm đã mỡ chưa
                invApp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
                // mở tài liệu part hiện có
                oPartDoc = (PartDocument)invApp.ActiveDocument;
                if (oPartDoc == null)
                {
                    MessageBox.Show("Không có part nào được mở");
                    return;
                }

                // lấy hệ thống tham số part
                Parameters parameters = oPartDoc.ComponentDefinition.Parameters;
                string ParameterName = parameterName;
                double NewValue;
                if (double.TryParse(newValue, out NewValue))
                {
                    UserParameter userParameter = parameters.UserParameters[ParameterName]; // nếu model thì đổi thành model
                    if (userParameter != null)
                    {
                        userParameter.Value = NewValue / 10;
                        oPartDoc.Update();
                        SetView();
                    }
                    else
                    {
                        MessageBox.Show(" lỗi không thể thay đổi tham số, vui lòng kiểm tra lại");
                    }
                }
                else
                {
                    MessageBox.Show("Không có giá trị nhập vào, vui lòng thêm giá trị vào");
                }
            }
            catch ( Exception ex)
            {
                MessageBox.Show(" lỗi chưa mở phần mềm, vui lòng mở phần mềm" + ex.Message);
            }
            
        }
        private void SetView()
        {
            try
            {
                Camera oCamera = invApp.ActiveView.Camera;
                oCamera.ViewOrientationType = ViewOrientationTypeEnum.kIsoTopRightViewOrientation;
                oCamera.Fit();
                oCamera.Apply();

            }
            catch (Exception ex)
            {
                MessageBox.Show("lỗi ko thể thiết lập góc nhìn" + ex.Message);
            }
        }
        private void ChangeMuitipleValueInPart(string parameterName1, string newValue1, string parameterName2, string newValue2, string parameterName3, string newValue3)
        {
            PartDocument oPartDoc = (PartDocument)invApp.ActiveDocument;
            if(oPartDoc == null)
            {
                MessageBox.Show("Không có tài liệu nào được mở");
                return;
            }

            // lấy hệ thống tham số trong part
            Parameters parameters = oPartDoc.ComponentDefinition.Parameters;
            bool updateRequired = false;

            // thay đổi tham số 1
            double NewValue1;
            if(double.TryParse(newValue1, out NewValue1))
            {
                UserParameter userParameter1 = parameters.UserParameters[parameterName1];
                if(userParameter1 != null)
                {
                    userParameter1.Value = NewValue1;
                    updateRequired = true;
                }
            }

            // thay đổi tham số 2
            double NewValue2;
            if (double.TryParse(newValue2, out NewValue2))
            {
                UserParameter userParameter2 = parameters.UserParameters[parameterName2];
                if (userParameter2 != null)
                {
                    userParameter2.Value = NewValue1;
                    updateRequired = true;
                }
            }

            // thay đổi tham số 3
            double NewValue3;
            if (double.TryParse(newValue3, out NewValue3))
            {
                UserParameter userParameter3 = parameters.UserParameters[parameterName3];
                if (userParameter3 != null)
                {
                    userParameter3.Value = NewValue1;
                    updateRequired = true;
                }
            }

            //cập nhật những thay đổi
            if (updateRequired) // cú pháp rút gọn cú pháp full updateRequired=true
            {
                oPartDoc.Update();
                MessageBox.Show("đã cập nhật thành công");
            }
            else
            {
                MessageBox.Show("không có giá trị được thay đổi");
            }

            SetView();
        }
        private void SuppressFeature(string featureName, bool suppress)
        {
            // xác nhận rằng tài liệu inventor đang chạy
            if (invApp == null) return;

            try
            {
                // lấy tài liệu hiện có 
                Document doc = invApp.ActiveDocument;
                if(doc.DocumentType != DocumentTypeEnum.kPartDocumentObject)
                {
                    MessageBox.Show("Không có tài liệu part nào đưuọc mở");
                    return;
                }

                PartDocument partDoc = (PartDocument)doc;

                // lấy định danh PartComponetDefinition
                PartComponentDefinition compDef = partDoc.ComponentDefinition;
                // lấy tất cả các tính năng (features) trong tài liệu
                PartFeatures features = compDef.Features;
                // tìm kiếm và bật/tắt tính năng
                bool success = false;

                foreach(ExtrudeFeature feature in features.ExtrudeFeatures)
                {
                    if(feature.Name == featureName)
                    {
                        feature.Suppressed = suppress;
                        success = true;
                        break;
                    }
                }

                if(! success)
                {
                    MessageBox.Show($"Feature'{featureName}' không tìm thấy");
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("error" + ex.Message);
            }
        }
        private void ConnectRunRuleIlogic(string ruleName)
        {
            if (invApp == null) return;

            try
            {
                //lấy tài liệu part hiện tại
                PartDocument oPartDoc = (PartDocument)invApp.ActiveDocument;
                if(oPartDoc == null || oPartDoc.DocumentType != DocumentTypeEnum.kPartDocumentObject)
                {
                    MessageBox.Show("Không có part nào được mở");
                    return;
                }

                // lấy định nghĩa PartComponentDefinition
                PartComponentDefinition comDef = oPartDoc.ComponentDefinition;
                // lấy Ilogic Automation add-in
                var addin = invApp.ApplicationAddIns.ItemById["{21E0AF7E-FBAB-4B25-B09B-0D0FA31A3DCE}"];
                if(addin == null)
                {
                    MessageBox.Show("add-in ilogic automation không có");
                    return;
                }

                dynamic ruleAutomation = addin.Automation;
                // chạy rule ilogic
                ruleAutomation.RunRule(oPartDoc, ruleName);
                MessageBox.Show($"chạy thành công mootk ilogic rule: {ruleName}");
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("lỗi chạy rule Ilogic rule:" + ex.Message);
            }
        }
        private void CreateInventorProject(string initialProjectPat,string nameProject)
        {
            try
            {
                //kiểm tra đưuòng dẫn không tồn tại thì tạo thư mục mới
                if(!System.IO.Directory.Exists(initialProjectPat))
                {
                    System.IO.Directory.CreateDirectory(initialProjectPat);
                }
                string projectFileName = nameProject +".ipj";
                string projectFilePatch = System.IO.Path.Combine(initialProjectPat, projectFileName);
                // lấy đối tượng DesignProjectManager
                DesignProjectManager designProjectManager = invApp.DesignProjectManager;
                //kiểm tra nếu dự án đã tồn tại
                if (System.IO.File.Exists(projectFilePatch))
                {
                    MessageBox.Show("Dự án đã tồn tại. Vui lòng nhập tên dự án khác");
                    return;
                }
                else
                {
                    //nếu dụ án chưa tồn tại thì taok sự án mới
                    DesignProject project = designProjectManager.DesignProjects.Add(MultiUserModeEnum.kSingleUserMode,projectFileName, projectFilePatch);
                    // mở dự án
                    project.Activate(true);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("lỗi khi tạo dự án" + ex.Message);
            }
        }

        //viết tổng quát lên

        public void ToggleFeatureSuppression(string featureName)
        {
            // Xác nhận rằng Inventor đang chạy
            if (invApp == null) return;

            try
            {
                // Lấy tài liệu Part hiện tại
                Document doc = invApp.ActiveDocument;

                if (doc.DocumentType != DocumentTypeEnum.kPartDocumentObject)
                {
                    MessageBox.Show("Active document is not a part document.");
                    return;
                }

                PartDocument partDoc = (PartDocument)doc;
                PartComponentDefinition compDef = partDoc.ComponentDefinition;

                // Tìm kiếm và tắt/bật (suppress/unsuppress) tính năng
                bool success = ToggleFeature(compDef.Features.ExtrudeFeatures, featureName);
                if (!success) success = ToggleFeature(compDef.Features.SweepFeatures, featureName);
                if (!success) success = ToggleFeature(compDef.Features.RevolveFeatures, featureName);
                if (!success) success = ToggleFeature(compDef.Features.LoftFeatures, featureName);
                if (!success) success = ToggleFeature(compDef.Features.HoleFeatures, featureName);
                if (!success) success = ToggleFeature(compDef.Features.FilletFeatures, featureName);
                if (!success) success = ToggleFeature(compDef.Features.ChamferFeatures, featureName);

                if (!success)
                {
                    MessageBox.Show($"Feature '{featureName}' not found.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        bool ToggleFeature(dynamic features, string nameToToggle)
        {
            for (int i = 1; i <= features.Count; i++)
            {
                PartFeature feature = features[i] as PartFeature;
                if (feature != null && feature.Name == nameToToggle)
                {
                    feature.Suppressed = !feature.Suppressed;
                    return true;
                }
            }
            return false;
        }
        private void ToggleFeatureSuppression_Ver2(string featureName)
        {
            // Xác nhận rằng Inventor đang chạy
            if (invApp == null) return;

            try
            {
                // Lấy tài liệu Part hiện tại
                Document doc = invApp.ActiveDocument;

                if (doc.DocumentType != DocumentTypeEnum.kPartDocumentObject)
                {
                    MessageBox.Show("Active document is not a part document.");
                    return;
                }

                PartDocument partDoc = (PartDocument)doc;

                // Lấy định danh PartComponentDefinition
                PartComponentDefinition compDef = partDoc.ComponentDefinition;

                // Lấy tất cả các tính năng (features) trong tài liệu
                PartFeatures features = compDef.Features;

                // Tìm kiếm và tắt/bật (suppress/unsuppress) tính năng
                bool success = false;

                foreach (ExtrudeFeature feature in features.ExtrudeFeatures)
                {
                    if (feature.Name == featureName)
                    {
                        feature.Suppressed = !feature.Suppressed;
                        success = true;
                        break;
                    }
                }
                foreach (SweepFeature feature in features.SweepFeatures)
                {
                    if (feature.Name == featureName)
                    {
                        feature.Suppressed = !feature.Suppressed;
                        success = true;
                        break;
                    }
                }
                if (!success)
                {
                    MessageBox.Show($"Feature '{featureName}' not found.");
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        //Pitch cạnh - xem xét lại
        private void PitchEdgeIfPartOpen(double pitchAngle)
        {
            if (invApp == null)
            {
                MessageBox.Show("Inventor is not running.");
                return;
            }

            try
            {
                // Lấy tài liệu Part hiện tại
                Document doc = invApp.ActiveDocument;

                if (doc.DocumentType != DocumentTypeEnum.kPartDocumentObject)
                {
                    MessageBox.Show("Active document is not a part document.");
                    return;
                }

                // Kết nối tài liệu Part và định nghĩa của nó
                PartDocument partDoc = (PartDocument)doc;
                PartComponentDefinition compDef = partDoc.ComponentDefinition;

                // Tạo HighlightSet để làm sáng cạnh
                HighlightSet highlightSet = invApp.ActiveDocument.HighlightSets.Add();

                // Yêu cầu người dùng chọn một cạnh
                MessageBox.Show("Please select an edge to pitch. Press 'Esc' to finish.");

                // Bật trạng thái picking
                isPicking = true;

                // Chọn cạnh
                while (isPicking)
                {
                    foreach (SelectionFilterEnum filter in new[] { SelectionFilterEnum.kPartEdgeFilter })
                    {
                        try
                        {
                            // Cho phép người dùng chọn cạnh qua Pick tool
                            ObjectCollection selectedEdges = invApp.CommandManager.Pick(filter, "Select an edge to pitch") as ObjectCollection;

                            if (selectedEdges != null && selectedEdges.Count > 0)
                            {
                                foreach (Edge edge in selectedEdges)
                                {
                                    // Thêm cạnh vào HighlightSet để làm sáng cạnh
                                    highlightSet.AddItem(edge);
                                }
                            }
                        }
                        catch (COMException)
                        {
                            // Handle pick tool cancel or any other COM exception gracefully
                        }
                    }

                    // Kiểm tra nếu người dùng nhấn Esc để thoát
                    if (GetAsyncKeyState((int)Keys.Escape) != 0)
                    {
                        isPicking = false;
                        MessageBox.Show("Edge selection mode finished.");
                    }
                }

                // Loại bỏ highlight sau khi lựa chọn kết thúc
                highlightSet.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void PitchEdgeIfPartOpen1()
        {
            if (invApp == null)
            {
                MessageBox.Show("Inventor is not running.");
                return;
            }

            bool isPicking = false;

            try
            {
                // Lấy tài liệu Part hiện tại
                Document doc = invApp.ActiveDocument;

                if (doc.DocumentType != DocumentTypeEnum.kPartDocumentObject)
                {
                    MessageBox.Show("Active document is not a part document.");
                    return;
                }

                // Kết nối tài liệu Part và định nghĩa của nó
                PartDocument partDoc = (PartDocument)doc;
                PartComponentDefinition compDef = partDoc.ComponentDefinition;

                // Tạo HighlightSet để làm sáng cạnh
                HighlightSet highlightSet = invApp.ActiveDocument.HighlightSets.Add();
                highlightSet.Color = invApp.TransientObjects.CreateColor(255, 0, 0);

                // Yêu cầu người dùng chọn một cạnh
                MessageBox.Show("Please select an edge to pitch. Press 'Esc' to finish.");

                // Bật trạng thái picking
                isPicking = true;

                // Chuỗi lưu tên cạnh đã chọn
                string selectedEdgeNames = string.Empty;

                // Chọn cạnh
                while (isPicking)
                {
                    foreach (SelectionFilterEnum filter in new[] { SelectionFilterEnum.kPartEdgeFilter })
                    {
                        try
                        {
                            // Cho phép người dùng chọn cạnh qua Pick tool
                            ObjectCollection selectedEdges = invApp.CommandManager.Pick(filter, "Select an edge to pitch") as ObjectCollection;

                            if (selectedEdges != null && selectedEdges.Count > 0)
                            {
                                foreach (Edge edge in selectedEdges)
                                {
                                    // Thêm cạnh vào HighlightSet để làm sáng cạnh
                                    highlightSet.AddItem(edge);

                                    // Tạo nhận dạng duy nhất cho cạnh và lưu vào chuỗi
                                    string edgeIdentifier = $"Edge with CurveType: {edge.GeometryType} and StartPoint: ({edge.StartVertex.Point.X}, {edge.StartVertex.Point.Y}, {edge.StartVertex.Point.Z})";
                                    selectedEdgeNames += edgeIdentifier + "\n";
                                }
                            }
                        }
                        catch (COMException)
                        {
                            // Xử lý hủy trình chọn hoặc lỗi COM khác
                        }
                    }

                    // Kiểm tra nếu người dùng nhấn Esc để thoát
                    if (GetAsyncKeyState((int)Keys.Escape) != 0)
                    {
                        isPicking = false;
                        MessageBox.Show("Edge selection mode finished.");
                    }
                }

                // Hiển thị tên các cạnh đã chọn
                if (!string.IsNullOrEmpty(selectedEdgeNames))
                {
                    MessageBox.Show("Selected Edges:\n" + selectedEdgeNames);
                }
                else
                {
                    MessageBox.Show("No edges were selected.");
                }

                // Loại bỏ highlight sau khi lựa chọn kết thúc
                highlightSet.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        //Dành cho Assembly

        private void ChangeParameterInAssembly(string parameterName, string newValue)
        {
            try
            {
                AssemblyDocument asmDoc = (AssemblyDocument)invApp.ActiveDocument;
                if (asmDoc == null)
                {
                    MessageBox.Show("Không có assembly nào được mở");
                    return;
                }

                // Lấy hệ thống tham số assembly
                Parameters asmParameters = asmDoc.ComponentDefinition.Parameters;

                double newParamValue;
                if (double.TryParse(newValue, out newParamValue))
                {
                    UserParameter userParameter = asmParameters.UserParameters[parameterName];
                    if (userParameter != null)
                    {
                        userParameter.Value = newParamValue / 10;
                        asmDoc.Update();
                        SetView();
                    }
                    else
                    {
                        MessageBox.Show("Lỗi không thể thay đổi tham số, vui lòng kiểm tra lại");
                    }
                }
                else
                {
                    MessageBox.Show("Giá trị nhập vào không hợp lệ, vui lòng thêm giá trị hợp lệ");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
        }

        private void ChangeParametersInAssembly(Dictionary<string, double> parametersToUpdate)
        {
            try
            {
                AssemblyDocument asmDoc = (AssemblyDocument)invApp.ActiveDocument;
                if (asmDoc == null)
                {
                    MessageBox.Show("Không có assembly nào được mở");
                    return;
                }

                // Lấy hệ thống tham số assembly
                Parameters asmParameters = asmDoc.ComponentDefinition.Parameters;
                bool updated = false;

                foreach (var entry in parametersToUpdate)
                {
                    string parameterName = entry.Key;
                    double newValue = entry.Value;

                    UserParameter userParameter = asmParameters.UserParameters[parameterName];
                    if (userParameter != null)
                    {
                        userParameter.Value = newValue / 10;
                        updated = true;
                    }
                    else
                    {
                        MessageBox.Show($"Lỗi không thể thay đổi tham số {parameterName}, vui lòng kiểm tra lại");
                    }
                }

                if (updated)
                {
                    asmDoc.Update();
                    SetView();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }

            /* cách sử dụng trên hàm chính
             * // Thay đổi nhiều tham số cùng lúc
                var parametersToUpdate = new Dictionary<string, double>
                {
                    { "ParameterName1", 20.0 },
                    { "ParameterName2", 45.0 },
                    { "ParameterName3", 70.0 }
                };

                ChangeParametersInAssembly(parametersToUpdate);
             * */
        }

        private void ToggleSuppress(string componentName)
        {
            try
            {
                Document doc = invApp.ActiveDocument;

                if (doc == null)
                {
                    MessageBox.Show("Không có tài liệu nào được mở.");
                    return;
                }

                switch (doc.DocumentType)
                {
                    case DocumentTypeEnum.kAssemblyDocumentObject:
                        ToggleSuppressInAssembly((AssemblyDocument)doc, componentName);
                        break;
                    case DocumentTypeEnum.kPartDocumentObject:
                        ToggleSuppressInPart((PartDocument)doc);
                        break;
                    default:
                        MessageBox.Show("Tài liệu hiện tại không phải là Assembly hoặc Part.");
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
        }
        public static void ToggleSuppressInAssembly(AssemblyDocument asmDoc, string componentName)
        {
            try
            {
                AssemblyComponentDefinition asmCompDef = asmDoc.ComponentDefinition;
                bool found = false;

                // Duyệt qua tất cả các thành phần của Assembly
                foreach (ComponentOccurrence occurrence in asmCompDef.Occurrences)
                {
                    if (occurrence.Name.Equals(componentName, StringComparison.OrdinalIgnoreCase))
                    {
                        // Bật/tắt trạng thái suppress của thành phần bằng cách gọi phương thức Suppress
                        occurrence.Suppress(!occurrence.Suppressed);
                        found = true;
                        break;
                    }
                }

                // Hiển thị thông báo kết quả
                if (found)
                {
                    asmDoc.Update();
                    MessageBox.Show($"Đã {(found ? "bật/tắt trạng thái suppress" : "không tìm thấy")} cho thành phần '{componentName}'.");
                }
                else
                {
                    MessageBox.Show($"Không tìm thấy thành phần '{componentName}' trong assembly.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
        }
        public static void ToggleSuppressInPart(PartDocument partDoc)
        {
            try
            {
                PartComponentDefinition partCompDef = partDoc.ComponentDefinition;
                bool found = false;

                // Duyệt qua tất cả các tính năng (features) và bật/tắt trạng thái suppress
                foreach (PartFeature feature in partCompDef.Features)
                {
                    feature.Suppressed = !feature.Suppressed;
                    found = true;
                }

                // Hiển thị thông báo kết quả
                if (found)
                {
                    partDoc.Update();
                    MessageBox.Show($"Đã {(found ? "bật/tắt trạng thái suppress" : "không tìm thấy tính năng")} cho part document.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
        }

        /// <summary>
        /// chưa chạy được
        /// </summary>
        private void SelectObjectsAndCreateProject()
        {
            if (invApp == null)
            {
                MessageBox.Show("Inventor is not running.");
                return;
            }

            try
            {
                Document doc = invApp.ActiveDocument;
                if (doc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    MessageBox.Show("Tài liệu hiện tại không phải là Assembly.");
                    return;
                }

                AssemblyDocument asmDoc = (AssemblyDocument)doc;
                AssemblyComponentDefinition asmCompDef = asmDoc.ComponentDefinition;

                // Yêu cầu người dùng chọn các đối tượng
                MessageBox.Show("Please select objects in the assembly. Press 'Esc' to finish.");

                // Bật trạng thái chọn các đối tượng
                ObjectCollection selectedObjects = invApp.TransientObjects.CreateObjectCollection();
                bool isPicking = true;

                while (isPicking)
                {
                    try
                    {
                        // Sử dụng Pick tool để chọn các đối tượng
                        Object selectedObject = invApp.CommandManager.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter, "Chọn đối tượng để lưu");

                        if (selectedObject != null)
                        {
                            selectedObjects.Add(selectedObject);
                        }
                    }
                    catch (COMException)
                    {
                        // Xử lý khi hủy trình chọn hoặc lỗi COM khác
                    }

                    // Kiểm tra nếu người dùng nhấn Esc để thoát
                    if (GetAsyncKeyState((int)Keys.Escape) != 0)
                    {
                        isPicking = false;
                    }
                }

                // Kiểm tra nếu không có đối tượng nào được chọn
                if (selectedObjects.Count == 0)
                {
                    MessageBox.Show("Không có object nào được chọn.");
                    return;
                }

                // In ra tên các đối tượng được chọn trong console
                Console.WriteLine("Các đối tượng được chọn:");
                foreach (ComponentOccurrence occurrence in selectedObjects)
                {
                    Console.WriteLine(occurrence.Name);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
        }

    }

}
