using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ConsoleApp1
{
    class FirstFuncation_Solid
    {
        /*
        Để sử dụng API SolidWorks trong một dự án C# (Windows Forms, Console Application, hoặc WPF), bạn cần phải thêm tham chiếu (reference) tới các thư viện COM của SolidWorks. Dưới đây là các bước chi tiết để thêm thư viện API SolidWorks vào dự án C# của bạn:

        bước 1: Mở Visual Studio và tạo một dự án mới hoặc mở một dự án hiện tại.

        bước 2: Trong Solution Explorer, nhấn chuột phải vào References, sau đó chọn Add Reference....

        bước 3: Trong cửa sổ Reference Manager, chọn tab COM.

        bước 4: Trong danh sách các thư viện COM, tìm và chọn các thư viện sau (bạn có thể cần cuộn xuống hoặc tìm từ khóa "SolidWorks"):
        
        SolidWorks 20xx Type Library(trong đó 20xx là phiên bản SolidWorks mà bạn đang sử dụng, ví dụ: 2015, 2018, 2020, ...). Đây là thư viện chính của SolidWorks.
        SolidWorks Constant Type Library(nếu bạn cần sử dụng các hằng số định nghĩa trong SolidWorks API).
        Các thư viện cơ bản thường là:

        SolidWorks.Interop.sldworks(SolidWorks Type Library)
        SolidWorks.Interop.swconst(SolidWorks Constant Type Library)
        Chú ý rằng bạn có thể cần phải cài đặt SolidWorks trên máy của bạn để các mục này xuất hiện trong danh sách.

       bước 4: Chọn các thư viện tương ứng và nhấn OK để thêm chúng vào dự án của bạn.

       bước 5: Kiểm tra References trong Solution Explorer để chắc chắn rằng các thư viện đã được thêm vào.
        */

        static SldWorks swApp;
        static ModelDoc2 swModel;
        static object[] sameHoleDataAtPart;
        static string DLResultpath;
        static object[] CSData;
        static object[] WData;
        static Edge oSelectedCircularEdge;

        static void Main(string[] args)
        {
            automaticPositionBolt_iam();
        }

        static void automaticPositionBolt_iam()
        {
            swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            swModel = (ModelDoc2)swApp.ActiveDoc;

            if (swModel == null)
            {
                MessageBox.Show("không có tài liệu nào đang mở");
                return;
            }

            try
            {
                if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    MessageBox.Show("Vui lòng mở một tài liệu lắp ráp.");
                    return;
                }

                // Tạo một selection filter
                swSelMgr = (SelectionMgr)swModel.SelectionManager;
                int mark = 0;

                // Filter the edges and get the user selected edge.
                oSelectedCircularEdge = (Edge)swSelMgr.GetSelectedObject6(
                    (int)swSelectType_e.swSelEDGES, mark);

                if (oSelectedCircularEdge == null)
                {
                    MessageBox.Show("không có cạnh nào được lựa chọn");
                    return;
                }

                Face2[] faceArray = (Face2[])oSelectedCircularEdge.GetTwoAdjacentFaces2();
                Component2 targetComp = (Component2)oSelectedCircularEdge.GetComponent();

                // Lấy đối tượng SurfaceBody và Geometry
                var targetBody = (Surface)oSelectedCircularEdge.GetSurface();
                MathPoint circleCenterPoint1 = ((MathUtility)swApp.GetMathUtility()).CreatePoint(oSelectedCircularEdge.GetCurve());

                double circleCenterDia1 = oSelectedCircularEdge.GetRadius() * 2 * 1000; // lấy đường kính cạnh hình tròn đã chọn

                if (faceArray.Length != 2)
                {
                    MessageBox.Show("Số lượng mặt kết nối của cạnh tròn không phải là 2!");
                    return;
                }

                // Trích xuất các mặt hình trụ và mặt phẳng trước và sau cạnh hình tròn đã chọn
                MathPoint circleCenterPoint2 = null;
                Face2 selectCylinderFace = null;
                Face2 selectPlaneFace = null;

                foreach (Face2 face in faceArray)
                {
                    var surfType = face.GetSurface().IsCylinder();
                    if (surType)
                    {
                        selectCylinderFace = face;
                    }
                    else if (face.GetSurface().IsPlanar())
                    {
                        selectPlaneFace = face;
                    }
                    else if (face.GetSurface().IsConical())
                    {
                        MessageBox.Show("Hiện tại không hỗ trợ bộ đếm lỗ dạng hình nón");
                        return;
                    }
                }

                // Kiểm tra xem mặt hình trụ lấy ra từ cạnh hình tròn có phải là lỗ không
                string checkHoleResult = CheckHoleExtrude(selectCylinderFace);
                if (checkHoleResult != "Hole")
                {
                    MessageBox.Show("Hãy chọn một lỗ cho cạnh hình tròn đã chọn");
                    return;
                }

                // Trích xuất vectơ dài nhất của bề mặt hình trụ
                MathVector HoleLengthVector1;
                PlacePointCollection = new object[] { }; // Phương pháp để khởi tạo tùy thuộc vào logic của bạn
                HoleLengthVector1 = GetMaxDistanceVector(selectCylinderFace, selectPlaneFace, PlacePointCollection);

                double holeLength1 = HoleLengthVector1.GetLength() * 1000; // mm

                // dev sau

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        static string CheckHoleExtrude(Face2 oFace)
        {
            // Để viết thêm logic kiểm tra
            return "Hole"; // Mặc định trả về là "Hole"; phải điều chỉnh theo logic thực
        }

        static MathVector GetMaxDistanceVector(Face2 cylinderFace, Face2 planeFace, object[] pointCollection)
        {
            // Implement logic here
            var maxVector = ((MathUtility)swApp.GetMathUtility()).CreateVector(new double[] { 0, 0, 0 });
            return maxVector;
        }

    }
}
