using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace OOP_InAPI
{
    public class inventorDrawing
    {
        private Inventor.Application ivnApp;

        // Border create and insert_ tạo và chèn border

        public void CreateBorderDefinition()
        {
            // Set a reference to the drawing document._ đặt tham chiếu vào giá trị tài liệu bản vẽ
            // This assumes a drawing document is active._ kiểm tra tài liệu drawing đã mở
            DrawingDocument oDrawDoc = (DrawingDocument)ivnApp.ActiveDocument;

            // Create the new border definition._ Tạo định nghĩa border mới
            BorderDefinition oBorderDef = oDrawDoc.BorderDefinitions.Add("Sample Border");

            // Open the border definition's sketch for edit. _ mở skech của định nghĩa border để điều chỉnh
            DrawingSketch oSketch;
            oBorderDef.Edit(out oSketch);

            TransientGeometry oTG = ivnApp.TransientGeometry;

            // Use the functionality of the sketch to add geometry._ sử dụng chức năng của sketch để thâm hình học
            oSketch.SketchLines.AddAsTwoPointRectangle(oTG.CreatePoint2d(2, 2), oTG.CreatePoint2d(25.94, 19.59));
            oSketch.SketchLines.AddByTwoPoints(oTG.CreatePoint2d(0, 10.795), oTG.CreatePoint2d(2, 10.795));
            oSketch.SketchLines.AddByTwoPoints(oTG.CreatePoint2d(13.97, 0), oTG.CreatePoint2d(13.97, 2));
            oSketch.SketchLines.AddByTwoPoints(oTG.CreatePoint2d(25.94, 10.795), oTG.CreatePoint2d(27.94, 10.795));
            oSketch.SketchLines.AddByTwoPoints(oTG.CreatePoint2d(13.97, 19.59), oTG.CreatePoint2d(13.97, 21.59));

            // Add some text to the border._ thêm một số text vào border
            Inventor.TextBox oTextBox = oSketch.TextBoxes.AddFitted(oTG.CreatePoint2d(2, 1), "Here is a sample string");
            oTextBox.VerticalJustification = VerticalTextAlignmentEnum.kAlignTextMiddle;

            // Add some prompted text to the border._ thêm một số yêu cầu từ người dùng
            oTextBox = oSketch.TextBoxes.AddFitted(oTG.CreatePoint2d(2, 20.59), "Enter designers name:");
            oTextBox.VerticalJustification = VerticalTextAlignmentEnum.kAlignTextMiddle;

            oBorderDef.ExitEdit(true);
        }

        public void InsertCustomBorderOnSheet()
        {
            // Set a reference to the drawing document._ đặt tham chiếu vào giá trị tài liệu bản vẽ
            // This assumes a drawing document is active._ kiểm tra tài liệu drawing đã mở
            DrawingDocument oDrawDoc = (DrawingDocument)ivnApp.ActiveDocument;

            // Obtain a reference to the desired border definition._ lấy tham chiếu đến định nghĩa border mong muốn
            BorderDefinition oBorderDef = oDrawDoc.BorderDefinitions["Sample Border"];

            Sheet oSheet = oDrawDoc.ActiveSheet;

            // Check to see if the sheet already has a border and delete it if it does._Kiểm tra sheet đã có border mong muốn hay chưa và xóa nó nếu có
            if (oSheet.Border != null)
            {
                oSheet.Border.Delete();
            }

            // This border definition contains one prompted string input._ Định nghĩa border này chứa một chuỗi yêu cầu input
            // An array must be input that contains the strings for the prompted strings._ một mảng sẽ được nhập vào chứa các chuỗi cho các yêu cầu input
            string[] sPromptStrings = new string[1];
            sPromptStrings[0] = "This is the input for the prompted text.";

            // Add an instance of the border definition to the sheet._ Thêm một phiên bảng của định nghĩa border vào sheet
            Border oBorder = oSheet.AddBorder(oBorderDef, sPromptStrings);
        }

        public void PlotAllSheetsInDrawing()
        {
            // In tất cả các trang trong tài liệu bản vẽ
            // Lấy tài liệu đang hoạt động và kiểm tra xem nó có phải là tài liệu bản vẽ không
            if (ivnApp.ActiveDocument.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
            {
                DrawingDocument oDrgDoc = (DrawingDocument)ivnApp.ActiveDocument;

                // Đặt tham chiếu vào trình quản lý in bản vẽ
                // DrawingPrintManager có nhiều tùy chọn hơn PrintManager
                // vì nó cụ thể cho tài liệu bản vẽ
                DrawingPrintManager oDrgPrintMgr = (DrawingPrintManager)oDrgDoc.PrintManager;

                // Đặt tên máy in
                // Bỏ qua dòng này để sử dụng máy in mặc định hoặc gán một máy in khác
                oDrgPrintMgr.Printer = "HP LaserJet 4000 Series PCL 6";

                // Đặt kích thước giấy, tỉ lệ và hướng
                oDrgPrintMgr.ScaleMode = PrintScaleModeEnum.kPrintBestFitScale;
                oDrgPrintMgr.PaperSize = PaperSizeEnum.kPaperSizeA4;
                oDrgPrintMgr.PrintRange = PrintRangeEnum.kPrintAllSheets;
                oDrgPrintMgr.Orientation = PrintOrientationEnum.kLandscapeOrientation;
                oDrgPrintMgr.SubmitPrint();
            }
        }

        public void PlotAllSheetsToPDF(string outputPdfPath)
        {
            // Xuất tất cả các trang trong tài liệu bản vẽ ra PDF
            // Lấy tài liệu đang hoạt động và kiểm tra xem nó có phải là tài liệu bản vẽ không
            if (ivnApp.ActiveDocument.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
            {
                DrawingDocument oDrgDoc = (DrawingDocument)ivnApp.ActiveDocument;

                // Tạo đối tượng TranslationContext
                TranslationContext oContext = ivnApp.TransientObjects.CreateTranslationContext();
                oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism;

                // Tạo đối tượng NameValueMap để lưu các tùy chọn xuất file
                NameValueMap oOptions = ivnApp.TransientObjects.CreateNameValueMap();

                // Lấy AddIn dịch PDF
                TranslatorAddIn PDFAddIn = null;
                foreach (ApplicationAddIn AddIn in ivnApp.ApplicationAddIns)
                {
                    if (AddIn.ClassIdString == "{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")
                    {
                        PDFAddIn = (TranslatorAddIn)AddIn;
                        break;
                    }
                }

                // Kiểm tra xem PDFAddIn có tồn tại hay không
                if (PDFAddIn != null)
                {
                    // Đặt tùy chọn xuất PDF
                    oOptions.set_Value("All_Color_AS_Black", 1);
                    oOptions.set_Value("Remove_Line_Weights", 0);
                    oOptions.set_Value("Vector_Resolution", 400);
                    oOptions.set_Value("Sheet_Range", PrintRangeEnum.kPrintAllSheets.ToString());

                    DataMedium oDataMedium = ivnApp.TransientObjects.CreateDataMedium();
                    oDataMedium.FileName = outputPdfPath;

                    // Xuất ra PDF
                    PDFAddIn.SaveCopyAs(oDrgDoc, oContext, oOptions, oDataMedium);
                }
                else
                {
                    Console.WriteLine("Cannot find PDF translator add-in or it does not support save copy as.");
                }
            }
            else
            {
                Console.WriteLine("Active document is not a drawing document.");
            }
        }

        public void InsertDefaultBorder()
        {
            // Đặt tham chiếu vào tài liệu bản vẽ
            // Điều này giả định rằng một tài liệu bản vẽ đang hoạt động
            DrawingDocument oDrawDoc = (DrawingDocument)ivnApp.ActiveDocument;

            // Đặt tham chiếu vào sheet hiện tại
            Sheet oSheet = oDrawDoc.ActiveSheet;

            // Kiểm tra xem sheet đã có border hay chưa và xóa nó nếu có
            if (oSheet.Border != null)
            {
                oSheet.Border.Delete();
            }

            // Định nghĩa các giá trị để sử dụng làm đầu vào cho việc tạo border
            int horizontalZoneCount = 15;
            BorderLabelModeEnum horizontalZoneLabelMode = BorderLabelModeEnum.kBorderLabelModeNumeric;
            int verticalZoneCount = 10;
            BorderLabelModeEnum verticalZoneLabelMode = BorderLabelModeEnum.kBorderLabelModeAlphabetical;
            bool labelFromBottomRight = false;
            bool delimitByLines = false;
            bool centerMarks = false;
            double topMargin = 5.0;
            double bottomMargin = 3.0;
            double leftMargin = 1.0;
            double rightMargin = 2.0;
            double borderLineWidth = 0.1;
            double textLabelHeight = 1.5;
            string font = "Courier New";

            // Thêm border vào sheet với các giá trị đã định nghĩa
            DefaultBorder oBorder = oSheet.AddDefaultBorder(
                horizontalZoneCount,
                horizontalZoneLabelMode,
                verticalZoneCount,
                verticalZoneLabelMode,
                labelFromBottomRight,
                delimitByLines,
                centerMarks,
                topMargin,
                bottomMargin,
                leftMargin,
                rightMargin,
                borderLineWidth,
                textLabelHeight,
                font);
        }
        
        public void InsertDefaultBorderUsingDefaults()
        {
            DrawingDocument oDrawDoc = (DrawingDocument)ivnApp.ActiveDocument;
            Sheet oSheet = oDrawDoc.ActiveSheet;

            if (oSheet.Border != null)
            {
                oSheet.Border.Delete();
            }

            DefaultBorder oBorder = oSheet.AddDefaultBorder();
        }

        public void CreateDrawingWithViews(string partFilePath, string drawingFilePath)
        {
            // Mở file Part
            PartDocument oPartDoc = (PartDocument)ivnApp.Documents.Open(partFilePath, true);

            // Tạo tài liệu bản vẽ mới
            DrawingDocument oDrawDoc = (DrawingDocument)ivnApp.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject);

            // Tạo một sheet mới
            Sheet oSheet = oDrawDoc.Sheets.Add(DrawingSheetSizeEnum.kA3DrawingSheetSize, PageOrientationTypeEnum.kLandscapePageOrientation);

            // Tạo một Base View
            TransientGeometry oTG = ivnApp.TransientGeometry;
            DrawingView oBaseView = oSheet.DrawingViews.AddBaseView(oPartDoc as _Document,
                oTG.CreatePoint2d(10, 10),
                1, ViewOrientationTypeEnum.kFrontViewOrientation,
                DrawingViewStyleEnum.kHiddenLineDrawingViewStyle);

            // Tạo một Projected View từ Base View
            DrawingView oProjView = oSheet.DrawingViews.AddProjectedView(oBaseView,
                oTG.CreatePoint2d(20, 10),
                DrawingViewStyleEnum.kHiddenLineDrawingViewStyle);

            // Tạo một Section View từ Projected View
            /*
            DrawingView oSectionView = oSheet.DrawingViews.AddSectionView(oProjView,
                oTG.CreatePoint2d(10, 15),
                DrawingSectionTypeEnum.kFullSectionDrawingSection,
                "A-A");
                */

            //DrawingView oSectionView = oSheet.DrawingViews.AddSectionView(oProjView,oTG.CreatePoint2d(10,15),DrawingViewStyleEnum.kFromBaseDrawingViewStyle,1,true,"A-A",true,true,true,);
            // Lưu tài liệu bản vẽ
            oDrawDoc.SaveAs(drawingFilePath, true);
        }

        public void CreateDrawingWithSheet(string drawingFilePath)
        {
            // Tạo tài liệu bản vẽ mới
            DrawingDocument oDrawDoc = (DrawingDocument)ivnApp.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject);

            // Tạo một sheet mới với kích thước và bố trí xác định
            Sheet osheet = oDrawDoc.Sheets.Add(DrawingSheetSizeEnum.kA3DrawingSheetSize, PageOrientationTypeEnum.kLandscapePageOrientation, "sheet");

            // Lưu tài liệu bản vẽ
            oDrawDoc.SaveAs(drawingFilePath, true);
        }



    }
}
