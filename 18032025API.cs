
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



private static void ImportAndMateParts(Inventor.Application inventorApp, string partFilePath, string boltFilePath)
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

                // Tạo lắp ghép trùng tâm của đường tâm lỗ và bu lông
                Console.WriteLine("Tạo lắp ghép đồng tâm...");
                asmCompDef.Constraints.AddInsertConstraint(holeEdge, boltAxis, true, 0);

                // Xác minh rằng cả hai bề mặt đều là mặt phẳng trước khi tạo lắp ghép bề mặt
                if (partFace.SurfaceType == SurfaceTypeEnum.kPlaneSurface && boltFace.SurfaceType == SurfaceTypeEnum.kPlaneSurface)
                {
                    // Tạo lắp ghép bề mặt giữa mặt của part và mặt của bu lông
                    Console.WriteLine("Tạo lắp ghép bề mặt...");
                    asmCompDef.Constraints.AddMateConstraint(partFace, boltFace, 0);
                }
                else
                {
                    Console.WriteLine("Một trong các bề mặt không phải là bề mặt phẳng, không thể tạo lắp ghép mặt phẳng.");
                }

                Console.WriteLine("Hoàn thành lắp ghép.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during assembly process: {ex.Message}");
            }
        }


 private void ImportAndMatePartsMultiple(Inventor.Application inventorApp, string partFilePath, string boltFilePath)
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

                // Lặp lại quá trình thêm bu lông và tạo đồng tâm lỗ cho đến khi người dùng nhấn nút Exit
                while (true)
                {
                    // Chọn tâm lỗ của chi tiết import
                    Console.WriteLine("Hãy chọn tâm lỗ của chi tiết...");
                    Edge holeEdge = inventorApp.CommandManager.Pick(SelectionFilterEnum.kPartEdgeFilter, "Chọn trục của lỗ.") as Edge;
                    if (holeEdge == null)
                    {
                        Console.WriteLine("Không chọn được trục.");
                        break;
                    }

                    // Chọn bề mặt của chi tiết import
                    Console.WriteLine("Hãy chọn bề mặt của chi tiết...");
                    Face partFace = inventorApp.CommandManager.Pick(SelectionFilterEnum.kPartFaceFilter, "Chọn bề mặt của chi tiết.") as Face;
                    if (partFace == null)
                    {
                        Console.WriteLine("Không chọn được bề mặt.");
                        break;
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
                        break;
                    }

                    // Chọn bề mặt của bu lông
                    Console.WriteLine("Hãy chọn bề mặt của bu lông...");
                    Face boltFace = inventorApp.CommandManager.Pick(SelectionFilterEnum.kPartFaceFilter, "Chọn bề mặt của bu lông.") as Face;
                    if (boltFace == null)
                    {
                        Console.WriteLine("Không chọn được bề mặt của bu lông.");
                        break;
                    }

                    // Tạo lắp ghép trùng tâm của đường tâm lỗ và bu lông
                    Console.WriteLine("Tạo lắp ghép đồng tâm...");
                    asmCompDef.Constraints.AddInsertConstraint(holeEdge, boltAxis, true, 0);

                    // Xác minh rằng cả hai bề mặt đều là mặt phẳng trước khi tạo lắp ghép bề mặt
                    if (partFace.SurfaceType == SurfaceTypeEnum.kPlaneSurface && boltFace.SurfaceType == SurfaceTypeEnum.kPlaneSurface)
                    {
                        // Tạo lắp ghép bề mặt giữa mặt của part và mặt của bu lông
                        Console.WriteLine("Tạo lắp ghép bề mặt...");
                        asmCompDef.Constraints.AddMateConstraint(partFace, boltFace, 0);
                    }
                    else
                    {
                        Console.WriteLine("Một trong các bề mặt không phải là bề mặt phẳng, không thể tạo lắp ghép mặt phẳng.");
                    }

                    // Kiểm tra nếu người dùng muốn thoát
                    Console.WriteLine("Press 'Enter' to insert another bolt or type 'exit' to quit:");
                    string userInput = Console.ReadLine();
                    if (userInput.ToLower() == "exit")
                    {
                        break;
                    }
                }

                Console.WriteLine("Hoàn thành lắp ghép.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during assembly process: {ex.Message}");
            }
        }

private static void ImportAndMatePartsLoop(Inventor.Application inventorApp, string partFilePath, string boltFilePath)
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

                // Lặp lại quá trình thêm bu lông và tạo đồng tâm lỗ cho đến khi người dùng nhấn phím ESC
                while (true)
                {
                    // Kiểm tra nếu người dùng nhấn ESC để thoát
                    if (IsKeyPressed(Keys.Escape))
                    {
                        Console.WriteLine("Đã nhấn phím ESC, thoát vòng lặp.");
                        break;
                    }

                    // Chọn tâm lỗ của chi tiết import
                    Console.WriteLine("Hãy chọn tâm lỗ của chi tiết...");
                    Edge holeEdge = inventorApp.CommandManager.Pick(SelectionFilterEnum.kPartEdgeFilter, "Chọn trục của lỗ.") as Edge;
                    if (holeEdge == null)
                    {
                        Console.WriteLine("Không chọn được trục.");
                        break;
                    }

                    // Chọn bề mặt của chi tiết import
                    Console.WriteLine("Hãy chọn bề mặt của chi tiết...");
                    Face partFace = inventorApp.CommandManager.Pick(SelectionFilterEnum.kPartFaceFilter, "Chọn bề mặt của chi tiết.") as Face;
                    if (partFace == null)
                    {
                        Console.WriteLine("Không chọn được bề mặt.");
                        break;
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
                        break;
                    }

                    // Chọn bề mặt của bu lông
                    Console.WriteLine("Hãy chọn bề mặt của bu lông...");
                    Face boltFace = inventorApp.CommandManager.Pick(SelectionFilterEnum.kPartFaceFilter, "Chọn bề mặt của bu lông.") as Face;
                    if (boltFace == null)
                    {
                        Console.WriteLine("Không chọn được bề mặt của bu lông.");
                        break;
                    }

                    // Tạo lắp ghép trùng tâm của đường tâm lỗ và bu lông
                    Console.WriteLine("Tạo lắp ghép đồng tâm...");
                    asmCompDef.Constraints.AddInsertConstraint(holeEdge, boltAxis, true, 0);

                    // Xác minh rằng cả hai bề mặt đều là mặt phẳng trước khi tạo lắp ghép bề mặt
                    if (partFace.SurfaceType == SurfaceTypeEnum.kPlaneSurface && boltFace.SurfaceType == SurfaceTypeEnum.kPlaneSurface)
                    {
                        // Tạo lắp ghép bề mặt giữa mặt của part và mặt của bu lông
                        Console.WriteLine("Tạo lắp ghép bề mặt...");
                        asmCompDef.Constraints.AddMateConstraint(partFace, boltFace, 0);
                    }
                    else
                    {
                        Console.WriteLine("Một trong các bề mặt không phải là bề mặt phẳng, không thể tạo lắp ghép mặt phẳng.");
                    }
                }

                Console.WriteLine("Hoàn thành lắp ghép.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during assembly process: {ex.Message}");
            }
        }

        private static bool IsKeyPressed(Keys key)
        {
            return (Control.ModifierKeys & key) == key;
        }


////////////////pitch object chạy theo con chuột chưa chạy được, đang lỗi

        /*
         *  private InteractionEvents interactionEvents;
        private ComponentOccurrence boltOcc;
        private AssemblyComponentDefinition asmCompDef;
        private Edge holeEdge;
        private Face partFace;
         * */
         /*
        private void ImportAndMatePartsWithInteraction(Inventor.Application inventorApp, string partFilePath, string boltFilePath)
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
                asmCompDef = asmDoc.ComponentDefinition;
                ComponentOccurrence partOcc = asmCompDef.Occurrences.Add(partFilePath, inventorApp.TransientGeometry.CreateMatrix());

                // Chọn tâm lỗ của chi tiết import
                Console.WriteLine("Hãy chọn tâm lỗ của chi tiết...");
                holeEdge = inventorApp.CommandManager.Pick(SelectionFilterEnum.kPartEdgeFilter, "Chọn trục của lỗ.") as Edge;
                if (holeEdge == null)
                {
                    Console.WriteLine("Không chọn được trục.");
                    return;
                }

                // Chọn bề mặt của chi tiết import
                Console.WriteLine("Hãy chọn bề mặt của chi tiết...");
                partFace = inventorApp.CommandManager.Pick(SelectionFilterEnum.kPartFaceFilter, "Chọn bề mặt của chi tiết.") as Face;
                if (partFace == null)
                {
                    Console.WriteLine("Không chọn được bề mặt.");
                    return;
                }

                // Thêm bu lông vào Assembly và di chuyển đến vị trí con chuột
                Console.WriteLine("Thêm bu lông vào Assembly...");
                boltOcc = asmCompDef.Occurrences.Add(boltFilePath, inventorApp.TransientGeometry.CreateMatrix());

                interactionEvents = inventorApp.CommandManager.CreateInteractionEvents();
                interactionEvents.OnTerminate += new InteractionEventsSink_OnTerminateEventHandler(InteractionEvents_OnTerminate);
                interactionEvents.OnMouseMove += new InteractionEventsSink_OnMouseMoveEventHandler(InteractionEvents_OnMouseMove);
                interactionEvents.OnSelect += new SelectEventsSink_OnSelectEventHandler(InteractionEvents_OnSelect);

                interactionEvents.StatusBarText = "Di chuyển bu lông đến vị trí và nhấn chọn để chèn.";
                interactionEvents.MouseMoveEnabled = true;
                interactionEvents.SelectEvents.Enabled = true;

                interactionEvents.Start();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during assembly process: {ex.Message}");
            }
        }

        private void InteractionEvents_OnSelect(Inventor.Application inventorApp, ObjectsEnumerator JustSelectedEntities, SelectionDeviceEnum SelectionDevice, Point ModelPosition, Point2d ViewPosition, Inventor.View View)
        {
            try
            {
                if (JustSelectedEntities != null && JustSelectedEntities.Count > 0)
                {
                    if (JustSelectedEntities[1] is Face selectedFace)
                    {
                        // Thiết lập vị trí của bu lông theo vị trí chuột
                        Matrix position = inventorApp.TransientGeometry.CreateMatrix();
                        position.SetTranslation(inventorApp.TransientGeometry.CreateVector(ModelPosition.X, ModelPosition.Y, ModelPosition.Z));
                        boltOcc.Transformation = position;

                        Face boltFace = null;
                        Edge boltAxis = null;

                        foreach (Face face in boltOcc.SurfaceBodies[1].Faces)
                        {
                            if (face.SurfaceType == SurfaceTypeEnum.kPlaneSurface)
                            {
                                boltFace = face;
                                break;
                            }
                        }

                        foreach (Edge edge in boltOcc.SurfaceBodies[1].Edges)
                        {
                            if (edge.CurveType == CurveTypeEnum.kCircularArcCurve)
                            {
                                boltAxis = edge;
                                break;
                            }
                        }

                        if (boltAxis != null && boltFace != null)
                        {
                            interactionEvents.Stop();
                            // Tạo lắp ghép trùng tâm của đường tâm lỗ và bu lông
                            Console.WriteLine("Tạo lắp ghép đồng tâm...");
                            asmCompDef.Constraints.AddInsertConstraint(holeEdge, boltAxis, true, 0);
                            // Tạo lắp ghép bề mặt giữa mặt của part và mặt của bu lông
                            Console.WriteLine("Tạo lắp ghép bề mặt...");
                            asmCompDef.Constraints.AddMateConstraint(partFace, boltFace, 0);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during selection: {ex.Message}");
            }
        }

        private void InteractionEvents_OnMouseMove(Inventor.Application inventorApp, MouseEventsData MouseEvents)
        {
            try
            {
                Matrix position = inventorApp.TransientGeometry.CreateMatrix();
                position.SetTranslation(inventorApp.TransientGeometry.CreateVector(MouseEvents.ModelPosition.X, MouseEvents.ModelPosition.Y, MouseEvents.ModelPosition.Z));
                boltOcc.Transformation = position;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during mouse move: {ex.Message}");
            }
        }

        private void InteractionEvents_OnTerminate()
        {
            try
            {
                // Giải phóng tài nguyên
                if (interactionEvents != null)
                {
                    Marshal.ReleaseComObject(interactionEvents);
                    interactionEvents = null;
                }
                if (boltOcc != null)
                {
                    Marshal.ReleaseComObject(boltOcc);
                    boltOcc = null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during termination: {ex.Message}");
            }
        }
        */
