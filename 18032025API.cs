
//ChangeMaterialToPart();
            //ChangeFaceColor(invApp);
            //ListAllRenderStyles(invApp);
            //SetOccurrenceRenderStyle(invApp);

            //string partFilePath = @"D:\1.OE_Keson\6.2025\6.Mechanical DE\0.Machine\1.ipt";
            //string boltFilePath = @"D:\1.OE_Keson\6.2025\6.Mechanical DE\0.Machine\2.ipt";


            //ImportAndAssembleParts(invApp, @"D:\1. OE_Keson\6.2025\6.Mechanical DE\0.Machine 1\1.ipt", @"D:\1. OE_Keson\6.2025\6.Mechanical DE\0.Machine 1\2.ipt");
            ImportPartAndBolt(invApp, @"D:\1. OE_Keson\6.2025\6.Mechanical DE\0.Machine 1\1.ipt", @"D:\1. OE_Keson\6.2025\6.Mechanical DE\0.Machine 1\2.ipt");
            //OpenFile(invApp, @"D:\1. OE_Keson\6.2025\6.Mechanical DE\0.Machine 1\2.ipt");






 /// <summary>
        /// 18/03/2025
        /// 1.làm việc với màu vật liệu và vật liệu
        /// </summary>
        private void ChangeMaterialToPart()
        {
            try
            {
                if (invApp == null)
                {
                    throw new InvalidOperationException("ứng dụng inventor chưa được tạo");
                }

                // Lấy tài liệu hiện tại và kiểm tra hợp lệ
                if (!(invApp.ActiveDocument is PartDocument))
                {
                    throw new InvalidOperationException("Tài liệu đang mở không phải là PartDocument.");
                }

                PartDocument oDoc = (PartDocument)invApp.ActiveDocument;
                string materialName = "Copper";
                 //string materialName = "Steel"; 
                 //string materialName = "Stainless Steel";

                // Lấy tài liệu thành phần hiện tại
                PartComponentDefinition compDef = oDoc.ComponentDefinition;

                // Tìm vật liệu trong thư viện tài liệu
                Material material = null;
                foreach (Material mat in oDoc.Materials)
                {
                    if (mat.Name.Equals(materialName, StringComparison.OrdinalIgnoreCase))
                    {
                        material = mat;
                        break;
                    }
                }

                // Nếu không tìm thấy vật liệu, thông báo lỗi và thoát
                if (material == null)
                {
                    throw new InvalidOperationException($"Không tìm thấy vật liệu '{materialName}' trong thư viện vật liệu của tài liệu.");
                }

                // Đặt vật liệu cho tài liệu hiện tại
                compDef.Material = material;

                // Làm mới hiển thị vật liệu trong giao diện người dùng
                invApp.CommandManager.ControlDefinitions["AppUpdateFXCmd"].Execute();

                Console.WriteLine("Hoàn thành việc đặt vật liệu cho phần.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi đặt vật liệu: {ex.Message}");
            }

        }

        //chưa chạ được
        private static void ChangeColorToPart(Inventor.Application inventorApp)
        {
            //try
            //{
            //    if (inventorApp == null)
            //    {
            //        throw new InvalidOperationException("Inventor Application is not initialized.");
            //    }

            //    // Lấy tài liệu hiện tại và kiểm tra hợp lệ
            //    if (!(inventorApp.ActiveDocument is PartDocument oDoc))
            //    {
            //        throw new InvalidOperationException("Tài liệu đang mở không phải là PartDocument.");
            //    }

            //    string renderStyleName = "Red"; // Đổi tên này thành màu bạn muốn sử dụng

            //    // Lấy StylesManager từ tài liệu hiện tại
            //    StylesManager stylesManager = inventorApp.StylesManager;

            //    // Kiểm tra xem RenderStyle có sẵn trong StylesManager không
            //    RenderStyle renderStyle = null;
            //    foreach (RenderStyle style in stylesManager.RenderStyles)
            //    {
            //        if (style.Name.Equals(renderStyleName, StringComparison.OrdinalIgnoreCase))
            //        {
            //            renderStyle = style;
            //            break;
            //        }
            //    }

            //    if (renderStyle == null)
            //    {
            //        throw new InvalidOperationException($"Không tìm thấy màu '{renderStyleName}' trong Render Styles.");
            //    }

            //    // Áp dụng Render Style cho phần hiện tại
            //    PartComponentDefinition compDef = oDoc.ComponentDefinition;
            //    Material material = compDef.Material;
            //    material.RenderStyle = renderStyle;

            //    // Làm mới hiển thị Render Style trong giao diện người dùng
            //    inventorApp.CommandManager.ControlDefinitions["AppUpdateFXCmd"].Execute();

            //    Console.WriteLine("Hoàn thành việc đổi màu cho phần.");
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine($"Lỗi khi đổi màu: {ex.Message}");
            //}
        }


        // Thay đổi màu bề mặt với option picth chọn
        private void ChangeFaceColor(Inventor.Application inventorApp)
        {
            try
            {
                if (inventorApp == null)
                {
                    throw new InvalidOperationException("Inventor Application is not initialized.");
                }

                // Lấy tài liệu hiện tại và kiểm tra hợp lệ
                if (!(inventorApp.ActiveDocument is PartDocument oDoc))
                {
                    throw new InvalidOperationException("Tài liệu đang mở không phải là PartDocument.");
                }

                // Chọn một mặt
                Face oFace = inventorApp.CommandManager.Pick(SelectionFilterEnum.kPartFaceFilter, "Select an Item") as Face;
                if (oFace == null)
                {
                    Console.WriteLine("Không chọn được mặt nào.");
                    return;
                }

                // Tên màu Render Style
                string renderStyleName = "Green";

                // Lấy Render Style từ tài liệu hiện tại
                RenderStyle renderStyle = null;
                try
                {
                    renderStyle = oDoc.RenderStyles[renderStyleName];
                }
                catch (Exception)
                {
                    Console.WriteLine($"Không thể chọn Style '{renderStyleName}'.");
                    return;
                }

                // Áp dụng Render Style cho mặt đã chọn
                oFace.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, renderStyle);

                Console.WriteLine("Hoàn thành việc đổi màu cho mặt đã chọn.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi đổi màu: {ex.Message}");
            }
        }

        //List tất cả các tên màu có trong list màu
        private static void ListAllRenderStyles(Inventor.Application inventorApp)
        {
            if (inventorApp.ActiveDocument is AssemblyDocument asmDoc)
            {
                foreach (RenderStyle style in asmDoc.RenderStyles)
                {
                    Console.WriteLine(style.Name);
                }
            }
        }

        // thay đổi màu assembly
        private void SetOccurrenceRenderStyle(Inventor.Application inventorApp)
        {
            try
            {
                if (inventorApp == null)
                {
                    throw new InvalidOperationException("Inventor Application is not initialized.");
                }

                // Lấy tài liệu hiện tại và kiểm tra hợp lệ
                if (!(inventorApp.ActiveDocument is AssemblyDocument asmDoc))
                {
                    throw new InvalidOperationException("Tài liệu đang mở không phải là AssemblyDocument.");
                }

                // Chọn một ComponentOccurrence
                ComponentOccurrence occ = inventorApp.CommandManager.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter, "Select an occurrence.") as ComponentOccurrence;

                if (occ == null)
                {
                    Console.WriteLine("Không chọn được ComponentOccurrence nào.");
                    return;
                }

                // Tên RenderStyle cho màu xanh nước biển
                string renderStyleName = "Dark Green"; // Thay đổi bằng tên chính xác bạn tìm thấy

                // Lấy RenderStyle từ danh sách RenderStyles
                RenderStyle renderStyle = null;
                foreach (RenderStyle style in asmDoc.RenderStyles)
                {
                    if (style.Name.Equals(renderStyleName, StringComparison.OrdinalIgnoreCase))
                    {
                        renderStyle = style;
                        break;
                    }
                }

                if (renderStyle == null)
                {
                    throw new InvalidOperationException($"Không tìm thấy màu '{renderStyleName}' trong Render Styles.");
                }

                // Gán RenderStyle cho ComponentOccurrence
                occ.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, renderStyle);

                Console.WriteLine("Hoàn thành việc thiết lập RenderStyle cho ComponentOccurrence.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi thiết lập RenderStyle: {ex.Message}");
            }
        }

        // chạy chưa được
        //private void ImportAndAssembleParts(Inventor.Application inventorApp, string partFilePath, string boltFilePath)
        //{
        //    try
        //    {
        //        if (inventorApp == null)
        //        {
        //            throw new InvalidOperationException("Inventor Application is not initialized.");
        //        }

        //        // Kiểm tra nếu tài liệu hiện tại là AssemblyDocument
        //        if (!(inventorApp.ActiveDocument is AssemblyDocument asmDoc))
        //        {
        //            throw new InvalidOperationException("The current document is not an AssemblyDocument.");
        //        }

        //        // Thêm chi tiết vào Assembly
        //        Console.WriteLine("Thêm part vào assembly...");
        //        AssemblyComponentDefinition asmCompDef = asmDoc.ComponentDefinition;
        //        ComponentOccurrence partOcc = asmCompDef.Occurrences.Add(partFilePath, inventorApp.TransientGeometry.CreateMatrix());

        //        // Chọn đường tâm lỗ và bề mặt của chi tiết
        //        Console.WriteLine("Please select the axis of the hole...");
        //        Edge holeAxis = inventorApp.CommandManager.Pick(SelectionFilterEnum.kPartEdgeFilter, "Lựa chọn tâm lỗ") as Edge;
        //        if (holeAxis == null)
        //        {
        //            Console.WriteLine("No axis selected.");
        //            return;
        //        }

        //        Console.WriteLine("Please select the face of the part...");
        //        Face selectedFace = inventorApp.CommandManager.Pick(SelectionFilterEnum.kPartFaceFilter, "Lựa chọn bề mặt") as Face;
        //        if (selectedFace == null)
        //        {
        //            Console.WriteLine("No face selected.");
        //            return;
        //        }

        //        // Thêm bu lông vào Assembly
        //        Console.WriteLine("Adding bolt to assembly...");
        //        ComponentOccurrence boltOcc = asmCompDef.Occurrences.Add(boltFilePath, inventorApp.TransientGeometry.CreateMatrix());

        //        // Chọn cạnh và bề mặt của bu lông
        //        Edge boltAxis = boltOcc.SurfaceBodies[1].Edges[1]; // Điều chỉnh số mục lục theo thực tế của bạn
        //        Face boltFace = boltOcc.SurfaceBodies[1].Faces[1]; // Điều chỉnh số mục lục theo thực tế của bạn

        //        // Lắp ghép trùng tâm của đường tâm lỗ và bu lông
        //        Console.WriteLine("Creating mate constraint for axis...");
        //        asmCompDef.Constraints.AddMateConstraint(
        //            holeAxis,
        //            boltAxis,
        //            0 // Khoảng cách lắp ghép
        //        );

        //        // Lắp ghép bề mặt dưới của bu lông với bề mặt được chọn của chi tiết
        //        Console.WriteLine("Creating mate constraint for face...");
        //        asmCompDef.Constraints.AddMateConstraint(
        //            selectedFace,
        //            boltFace,
        //            0 // Khoảng cách lắp ghép
        //        );

        //        Console.WriteLine("Assembly completed successfully.");
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"Error during assembly process: {ex.Message}");
        //    }
        //}

        //kiểm tra mở một tài liệu part
        private Document OpenFile(Inventor.Application inventorApp, string filePath)
        {
            Document doc = null;
            try
            {
                if (inventorApp == null)
                {
                    throw new InvalidOperationException("Inventor Application is not initialized.");
                }

                if (System.IO.File.Exists(filePath))
                {
                    doc = inventorApp.Documents.Open(filePath, true);
                    Console.WriteLine($"Đã mở file: {filePath}");
                }
                else
                {
                    Console.WriteLine("File không tồn tại.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi mở file: {ex.Message}");
            }
            return doc;
        }

        private void ImportPartAndBolt(Inventor.Application inventorApp, string partFilePath, string boltFilePath)
        {
            try
            {
                // Kiểm tra nếu tài liệu hiện tại là AssemblyDocument
                if (!(inventorApp.ActiveDocument is AssemblyDocument asmDoc))
                {
                    throw new InvalidOperationException("The current document is not an AssemblyDocument.");
                }

                // Thêm chi tiết vào Assembly
                Console.WriteLine("Thêm phần tử vào Assembly...");
                AssemblyComponentDefinition asmCompDef = asmDoc.ComponentDefinition;
                ComponentOccurrence partOcc = asmCompDef.Occurrences.Add(partFilePath, inventorApp.TransientGeometry.CreateMatrix());

                // Chọn tâm lỗ của chi tiết import
                Console.WriteLine("Hãy chọn tâm lỗ của chi tiết...");
                Edge holeEdge = inventorApp.CommandManager.Pick(SelectionFilterEnum.kPartEdgeFilter, "Chọn trục của lỗ.") as Edge;
                if (holeEdge == null)
                {
                    Console.WriteLine("Không chọn được trục.");
                    return;
                }

                // Chọn bề mặt của chi tiết import
                Console.WriteLine("Hãy chọn bề mặt của chi tiết...");
                Face partFace = inventorApp.CommandManager.Pick(SelectionFilterEnum.kPartFaceFilter, "Chọn bề mặt của chi tiết.") as Face;
                if (partFace == null)
                {
                    Console.WriteLine("Không chọn được bề mặt.");
                    return;
                }

                // Thêm bu lông vào Assembly
                Console.WriteLine("Thêm bu lông vào Assembly...");
                ComponentOccurrence boltOcc = asmCompDef.Occurrences.Add(boltFilePath, inventorApp.TransientGeometry.CreateMatrix());

                // Chọn trục của bu lông
                Console.WriteLine("Hãy chọn trục của bu lông...");
                Edge boltAxis = inventorApp.CommandManager.Pick(SelectionFilterEnum.kPartEdgeFilter, "Chọn trục của bu lông.") as Edge;
                if (boltAxis == null)
                {
                    Console.WriteLine("Không chọn được trục của bu lông.");
                    return;
                }

                // Chọn bề mặt của bu lông
                Console.WriteLine("Hãy chọn bề mặt của bu lông...");
                Face boltFace = inventorApp.CommandManager.Pick(SelectionFilterEnum.kPartFaceFilter, "Chọn bề mặt của bu lông.") as Face;
                if (boltFace == null)
                {
                    Console.WriteLine("Không chọn được bề mặt của bu lông.");
                    return;
                }

                // Lắp ghép trùng tâm của đường tâm lỗ và bu lông
                Console.WriteLine("Tạo lắp ghép đồng tâm...");
                asmCompDef.Constraints.AddInsertConstraint(holeEdge, boltAxis, true, 0);

                // Lắp ghép bề mặt dưới của bu lông với bề mặt được chọn của chi tiết
                Console.WriteLine("Tạo lắp ghép bề mặt...");
                asmCompDef.Constraints.AddMateConstraint(partFace, boltFace, 0);

                Console.WriteLine("Hoàn thành lắp ghép.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during assembly process: {ex.Message}");
            }
        }
