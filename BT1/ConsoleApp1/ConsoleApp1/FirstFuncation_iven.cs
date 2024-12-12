using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ConsoleApp1
{
    class FirstFuncation_iven
    {

        /*   
            Để sử dụng API Autodesk Inventor trong dự án C# của bạn, bạn sẽ cần thêm tham chiếu (reference) đến thư viện COM của Autodesk Inventor. Dưới đây là các bước chi tiết để thêm thư viện API Inventor vào dự án C# của bạn:

            bước 1: Mở Visual Studio và tạo một dự án mới hoặc mở một dự án hiện tại.

            bước 2: Trong Solution Explorer, nhấn chuột phải vào References, sau đó chọn Add Reference....

            bước 3: Trong cửa sổ Reference Manager, chọn tab COM.

            bước 4: Trong danh sách các thư viện COM, tìm và chọn:

            Autodesk Inventor Object Library
            Bạn có thể cần phải cuộn xuống hoặc tìm kiếm từ khóa "Inventor" để tìm thấy thư viện này.

            bước 5: Chọn thư viện Autodesk Inventor Object Library và nhấn OK để thêm vào dự án của bạn.

            bước 6: Kiểm tra References trong Solution Explorer để chắc chắn rằng thư viện đã được thêm vào.
        */
        /// <summary>
        /// macro inport ốc tự động theo kích thước lỗ từ 1 thư viện cá nhân trong inventor
        /// </summary>

        static ObjectCollection PlacePointCollection;
        static ObjectCollection Face2Collection;
        static Face selectCylinderFace;
        static Face selectPlanFace;
        static object[] sameHoleDataAtPart;
        static string DLResultpath;
        static object[] CSData;
        static object[] WData;
        static Edge oSelectedCircularEdge;

        static void automaticPositionBolt_iam()
        {
            Inventor.Application ThisApplication = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
            Document ThisDoc = ThisApplication.ActiveDocument;

            if (ThisDoc == null)
            {
                MessageBox.Show("không có tài liệu nào đang mở");
            }

            try
            {
                Document OEDoc = ThisApplication.ActiveEditDocument;
                if (OEDoc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    MessageBox.Show(" đoạn comment ko biết ");
                }

                // Tạo một selection filter
                SelectionFilterEnum oSelectionFilter;
                oSelectionFilter = SelectionFilterEnum.kPartEdgeCircularFilter; // xem xét lại kCircularEdgeFilter nằm ở đâu 

                // Filter the edges and get the user selected edge.
                oSelectedCircularEdge = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kPartEdgeCircularFilter, "");

                if (oSelectedCircularEdge == null)
                {
                    MessageBox.Show("không có cạnh nào được lựa chọn");
                    return;
                }

                ComponentOccurrence targetOcc = oSelectedCircularEdge.Parent.Parent; // nhận sự xuất hiện mục tiêu
                SurfaceBody targetBody = oSelectedCircularEdge.Parent; // lấy chiều dài lỗ

                Point circleCenterPoint1 = oSelectedCircularEdge.Geometry.center; // lấy tọa độ tâm của cạnh tròn đã chọn
                double circleCenterDia1 = oSelectedCircularEdge.Geometry.Radius * 2 * 10; // lấy đường kính cạnh hình tròn đã chọn

                if (oSelectedCircularEdge.Faces.Count != 2)
                {
                    MessageBox.Show("Đây là trường hợp kết nối FaceCount của cạnh hình tròn không phải là 2!");
                }

                //Trích xuất các mặt hình trụ và mặt phẳng trước và sau cạnh hình tròn đã chọn
                Point circleCenterPoint2 = null;
                Face selectCylinderFace = null;
                Face selectPlaneFace = null;

                foreach (Face face1 in oSelectedCircularEdge.Faces)
                {
                    if (face1.SurfaceType == SurfaceTypeEnum.kCylinderSurface)
                    {
                        selectCylinderFace = face1;
                    }
                    else if (face1.SurfaceType == SurfaceTypeEnum.kPlaneSurface)
                    {
                        selectPlaneFace = face1;
                    }
                    else if (face1.SurfaceType == SurfaceTypeEnum.kConeSurface)
                    {
                        MessageBox.Show("bộ đếm không đưuọc hỗ trợ");
                        break; // thoát khỏi vòng lặp
                    }
                }

                //Kiểm tra xem hình trụ được lấy ra từ cạnh hình tròn đã chọn có phải là lỗ không
                string checkHoleResult = CheckHoleExtrude(selectCylinderFace);
                if (checkHoleExtrude != "Hole")
                {
                    MessageBox.Show("Hãy chọn một lỗ cho cạnh hình tròn đã chọn");
                    //thoát chưa nghĩ ra
                    //GoTo SubEnd // xem xét hàm chỗ này
                }

                //Trích xuất vectơ dài nhất của bề mặt hình trụ
                Vector HoleLengthVector1;
                ObjectCollection PlacePointCollection = ThisApplication.TransientObjects.CreateObjectCollection();
                HoleLengthVector1 = GetMaxDistanceVector(selectCylinderFace, selectPlaneFace, PlacePointCollection);

                double holeLength1 = HoleLengthVector1.Length * 10; // *10 chuyển qua mm

                // dev sau

            }
            catch
            {

            }
        }

        static string checkHoleExtrude;
        static string CheckHoleExtrude(Face oFace)
        {
            // để viết thêm chua đủ time để chuyển
            return checkHoleExtrude;
            //return null;
        }

        static Vector GetMaxDistanceVector(Face cylinderFace, Face planeFace, ObjectCollection pointCollection)
        {
            // Implement logic here
            // Trả về vector tương ứng
            return null;
        }
    }
}
