using System;
using System.Collections.Generic;


namespace OOP
{
    class Program
    {
        static void Main(string[] args)
        {
            // tạo danh sách các bệnh nhân
            List<Patient> patients = new List<Patient>();

            // thêm danh sách người lớn
            patients.Add(new Adult {Name = "A", Age = 30, Gender = "Male" });
            patients.Add(new Adult {Name = "B", Age = 36, Gender = "FeMale" }); // thêm cách thủ công

            // thêm danh sách trẻ e vào danh sách
            for ( int i =1; i<= 5; i++)
            {
                patients.Add(new Child {Name = $"child {i}", Age = 5 + i, Gender = "Male" });
            }

            //in tất cả các giá trị có trong danh sách
            //foreach ( var patient in patients)
            //{
            //    patient.PrintInfo();
            //}

            Console.WriteLine("Second patient in the list:");

            // truy xuất 1 giá trị của list theo vị trí
            if (patients.Count > 1)
            {
                patients[0].PrintInfo();
            }
            Console.ReadKey();
        }
    }

    // lớp cơ bản Patient
    public class Patient
    {
        public string Name { get; set;}
        public int Age { get; set; }
        public string Gender { get; set; }
        //Phương thức ảo để in thông tin bênh nhân

        public virtual void PrintInfo()
        {
            Console.WriteLine($"Patient: {Name}, Age : {Age}, Gender: {Gender}"); // ghi để truyền giá trị biến vào nhanh
        }
    }

    // lớp adult kế thừa từ bệnh nhân
    public class Adult:Patient
    {
        public override void PrintInfo()
        {
            Console.WriteLine($"Aldult: {Name}, Age : {Age}, Gender: {Gender}"); // ghi để truyền giá trị biến vào nhanh
        }
    }

    // lớp Child kế thừa từ bệnh nhân
    public class Child:Patient
    {
        public override void PrintInfo()
        {
            Console.WriteLine($"Child: {Name}, Age : {Age}, Gender: {Gender}"); // ghi để truyền giá trị biến vào nhanh
        }
    }
}
