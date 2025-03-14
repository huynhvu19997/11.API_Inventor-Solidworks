private void AnalyzeInterferenceInActiveAssembly()
        {
            try
            {
                if (invApp != null)
                {
                    Console.WriteLine("Kết nối thành công với Autodesk Inventor.");

                    // Lấy tài liệu Assembly đang mở
                    AssemblyDocument asmDoc = (AssemblyDocument)invApp.ActiveDocument;

                    if (asmDoc != null)
                    {
                        // Lấy ComponentDefinition của Assembly
                        AssemblyComponentDefinition asmCompDef = asmDoc.ComponentDefinition;

                        // Tạo danh sách các thành phần kiểm tra giao thoa
                        ObjectCollection allComponents = invApp.TransientObjects.CreateObjectCollection();
                        foreach (ComponentOccurrence comp in asmCompDef.Occurrences)
                        {
                            allComponents.Add(comp);
                        }

                        // Phân tích sự giao thoa
                        InterferenceResults interferenceResults = asmCompDef.AnalyzeInterference(allComponents, null);

                        // Kiểm tra và hiển thị kết quả
                        if (interferenceResults.Count > 0)
                        {
                            foreach (InterferenceResult result in interferenceResults)
                            {
                                Console.WriteLine($"Sự giao thoa giữa {result.OccurrenceOne.Name} và {result.OccurrenceTwo.Name}");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Không tìm thấy sự giao thoa nào.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Không có tài liệu assembly nào đang mở.");
                    }
                }
                else
                {
                    Console.WriteLine("Không thể kết nối với Autodesk Inventor.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Đã xảy ra lỗi: {ex.Message}");
                Console.WriteLine($"Chi tiết: {ex.StackTrace}");
            }
        }

        private  void AnalyzeInterferenceInActiveAssembly1()
        {
            try
            {
                if (invApp != null)
                {
                    // Lấy tài liệu Assembly đang mở
                    AssemblyDocument asmDoc = (AssemblyDocument)invApp.ActiveDocument;

                    if (asmDoc != null)
                    {
                        // Lấy ComponentDefinition của Assembly
                        AssemblyComponentDefinition asmCompDef = asmDoc.ComponentDefinition;

                        // Tạo danh sách các thành phần kiểm tra giao thoa
                        ObjectCollection allComponents = invApp.TransientObjects.CreateObjectCollection();
                        foreach (ComponentOccurrence comp in asmCompDef.Occurrences)
                        {
                            allComponents.Add(comp);
                        }

                        // Phân tích sự giao thoa
                        InterferenceResults interferenceResults = asmCompDef.AnalyzeInterference(allComponents, null);

                        // Kiểm tra và hiển thị kết quả
                        if (interferenceResults.Count > 0)
                        {
                            string message = "Sự giao thoa được phát hiện:\n";
                            foreach (InterferenceResult result in interferenceResults)
                            {
                                message += $"Sự giao thoa giữa {result.OccurrenceOne.Name} và {result.OccurrenceTwo.Name}\n";
                            }
                            MessageBox.Show(message, "Kết quả Giao thoa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("Không tìm thấy sự giao thoa nào.", "Kết quả Giao thoa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Không có tài liệu assembly nào đang mở.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Không thể kết nối với Autodesk Inventor.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Đã xảy ra lỗi: {ex.Message}\nChi tiết: {ex.StackTrace}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


//////////////
 private void AnalyzeInterferenceAndSuppressComponents()
        {
            try
            {
                if (invApp != null)
                {
                    Console.WriteLine("Kết nối thành công với Autodesk Inventor.");

                    // Lấy tài liệu Assembly đang mở
                    AssemblyDocument asmDoc = (AssemblyDocument)invApp.ActiveDocument;

                    if (asmDoc != null)
                    {
                        // Lấy ComponentDefinition của Assembly
                        AssemblyComponentDefinition asmCompDef = asmDoc.ComponentDefinition;

                        // Tạo danh sách các thành phần kiểm tra sự giao thoa
                        ObjectCollection allComponents = invApp.TransientObjects.CreateObjectCollection();
                        foreach (ComponentOccurrence comp in asmCompDef.Occurrences)
                        {
                            allComponents.Add(comp);
                        }

                        // Phân tích sự giao thoa
                        InterferenceResults interferenceResults = asmCompDef.AnalyzeInterference(allComponents);

                        // Kiểm tra kết quả và vô hiệu hóa các thành phần
                        if (interferenceResults.Count > 0)
                        {
                            foreach (InterferenceResult result in interferenceResults)
                            {
                                SuppressComponent(result.OccurrenceOne);
                                SuppressComponent(result.OccurrenceTwo);
                            }
                            Console.WriteLine("Các thành phần gây ra sự giao thoa đã được vô hiệu hóa.");
                        }
                        else
                        {
                            Console.WriteLine("Không tìm thấy sự giao thoa nào.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Không có tài liệu assembly nào đang mở.");
                    }
                }
                else
                {
                    Console.WriteLine("Không thể kết nối với Autodesk Inventor.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Đã xảy ra lỗi: {ex.Message}");
                Console.WriteLine($"Chi tiết: {ex.StackTrace}");
            }
        }

        private void SuppressComponent(ComponentOccurrence component)
        {
            try
            {
                if (!component.Suppressed)
                {
                    component.Suppress(true); // Vô hiệu hóa (suppress) thành phần
                    Console.WriteLine($"Component {component.Name} đã được vô hiệu hóa.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Đã xảy ra lỗi khi vô hiệu hóa component {component.Name}: {ex.Message}");
            }
        }
//////////////////kiểm tra và xuất list danh sách ra txt

private void AnalyzeInterferenceAndSuppressComponents1()
        {
            try
            {
                if (invApp != null)
                {
                    Console.WriteLine("Kết nối thành công với Autodesk Inventor.");

                    // Lấy tài liệu Assembly đang mở
                    AssemblyDocument asmDoc = (AssemblyDocument)invApp.ActiveDocument;

                    if (asmDoc != null)
                    {
                        // Lấy ComponentDefinition của Assembly
                        AssemblyComponentDefinition asmCompDef = asmDoc.ComponentDefinition;

                        // Tạo danh sách các thành phần kiểm tra sự giao thoa
                        ObjectCollection allComponents = invApp.TransientObjects.CreateObjectCollection();
                        foreach (ComponentOccurrence comp in asmCompDef.Occurrences)
                        {
                            allComponents.Add(comp);
                        }

                        // Phân tích sự giao thoa
                        InterferenceResults interferenceResults = asmCompDef.AnalyzeInterference(allComponents);

                        // Lưu danh sách các chi tiết va chạm ra file
                        string filePath = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), "InterferenceResults.txt");
                        using (StreamWriter writer = new StreamWriter(filePath))
                        {
                            if (interferenceResults.Count > 0)
                            {
                                writer.WriteLine("Danh sách các chi tiết va chạm:");
                                foreach (InterferenceResult result in interferenceResults)
                                {
                                    string message = $"Giao thoa giữa {result.OccurrenceOne.Name} và {result.OccurrenceTwo.Name}";
                                    writer.WriteLine(message);
                                    SuppressComponent(result.OccurrenceOne);
                                    SuppressComponent(result.OccurrenceTwo);
                                }
                                Console.WriteLine("Các thành phần gây ra sự giao thoa đã được vô hiệu hóa và danh sách được lưu ra file.");
                            }
                            else
                            {
                                writer.WriteLine("Không tìm thấy sự giao thoa nào.");
                                Console.WriteLine("Không tìm thấy sự giao thoa nào.");
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("Không có tài liệu assembly nào đang mở.");
                    }
                }
                else
                {
                    Console.WriteLine("Không thể kết nối với Autodesk Inventor.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Đã xảy ra lỗi: {ex.Message}");
                Console.WriteLine($"Chi tiết: {ex.StackTrace}");
            }
        }

        private  void SuppressComponent1(ComponentOccurrence component)
        {
            try
            {
                if (!component.Suppressed)
                {
                    component.Suppress(true); // Vô hiệu hóa (suppress) thành phần sử dụng đúng phương thức
                    Console.WriteLine($"Component {component.Name} đã được vô hiệu hóa.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Đã xảy ra lỗi khi vô hiệu hóa component {component.Name}: {ex.Message}");
            }
        }
//////////////lưu với đường dẫn linh hoạt

private  void AnalyzeInterferenceAndSuppressComponents2()
        {
            try
            {
                if (invApp != null)
                {
                    Console.WriteLine("Kết nối thành công với Autodesk Inventor.");

                    // Lấy tài liệu Assembly đang mở
                    AssemblyDocument asmDoc = (AssemblyDocument)invApp.ActiveDocument;

                    if (asmDoc != null)
                    {
                        // Lấy ComponentDefinition của Assembly
                        AssemblyComponentDefinition asmCompDef = asmDoc.ComponentDefinition;

                        // Tạo danh sách các thành phần kiểm tra sự giao thoa
                        ObjectCollection allComponents = invApp.TransientObjects.CreateObjectCollection();
                        foreach (ComponentOccurrence comp in asmCompDef.Occurrences)
                        {
                            allComponents.Add(comp);
                        }

                        // Phân tích sự giao thoa
                        InterferenceResults interferenceResults = asmCompDef.AnalyzeInterference(allComponents);

                        // Hiển thị hộp thoại lưu tệp
                        SaveFileDialog saveFileDialog = new SaveFileDialog();
                        saveFileDialog.Filter = "Text Files|*.txt";
                        saveFileDialog.Title = "Save Interference Results";
                        saveFileDialog.FileName = "InterferenceResults.txt";

                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            string filePath = saveFileDialog.FileName;
                            using (StreamWriter writer = new StreamWriter(filePath))
                            {
                                if (interferenceResults.Count > 0)
                                {
                                    writer.WriteLine("Danh sách các chi tiết va chạm:");
                                    foreach (InterferenceResult result in interferenceResults)
                                    {
                                        string message = $"Giao thoa giữa {result.OccurrenceOne.Name} và {result.OccurrenceTwo.Name}";
                                        writer.WriteLine(message);
                                        SuppressComponent(result.OccurrenceOne);
                                        SuppressComponent(result.OccurrenceTwo);
                                    }
                                    Console.WriteLine("Các thành phần gây ra sự giao thoa đã được vô hiệu hóa và danh sách được lưu ra file.");
                                }
                                else
                                {
                                    writer.WriteLine("Không tìm thấy sự giao thoa nào.");
                                    Console.WriteLine("Không tìm thấy sự giao thoa nào.");
                                }
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("Không có tài liệu assembly nào đang mở.");
                    }
                }
                else
                {
                    Console.WriteLine("Không thể kết nối với Autodesk Inventor.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Đã xảy ra lỗi: {ex.Message}");
                Console.WriteLine($"Chi tiết: {ex.StackTrace}");
            }
        }

        private  void SuppressComponent2(ComponentOccurrence component)
        {
            try
            {
                if (!component.Suppressed)
                {
                    component.Suppress(true); // Vô hiệu hóa (suppress) thành phần
                    Console.WriteLine($"Component {component.Name} đã được vô hiệu hóa.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Đã xảy ra lỗi khi vô hiệu hóa component {component.Name}: {ex.Message}");
            }
        }
