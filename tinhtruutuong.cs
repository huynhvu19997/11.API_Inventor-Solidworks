using System;
using System.Collections.Generic;


namespace OOP
{
    class Program
    {
        static void Main(string[] args)
        {
            // tạo đối tượng Circle và in diện tích
            shape circle = new Circle(5);
            circle.PrintArea();  // in diện tích hình tròn

            // tạo đối tượng rectangle và in diện tích
            shape rectangle = new Rectangle(4, 6);
            rectangle.PrintArea();
        }
    }

    // định nghĩa lớp trừu tượng Shape

    public abstract class shape
    {
        // phương thức trừu tượng tính diện tích
        public abstract double CalculateArea();
        // phương thức cụ thể in thông tin diện tích
        public void PrintArea()
        {
            Console.WriteLine($"khu vuc là {CalculateArea()}");
        }
    }

    //Circle kế thừa lớp shape
    public class Circle:shape
    {
        public double Radius { get; set; }

        public Circle(double radius) // hàm khởi tạo đối tượng vói bán kính dduocj cung cấp
        {
            Radius = radius;
        }

        // Thực hiện phương thức trừu tượng CalculateArea
        public override double CalculateArea() // ghi đè để tính hình tròn
        {
            return Math.PI * Radius * Radius;
        }
    }

    // lớp rectangle kế thừa từ lớp shape

    public class Rectangle : shape
    {
        public double _width { get; set; } // hoặc theo quy tắc Width
        public double _height { get; set; }

        public Rectangle(double width, double height) // biến truyền vào ko viết hoa
        {
            _width = width;
            _height = height;
        }

        // Thực hiện phương thức trừu tượng CalculateArea
        public override double CalculateArea() // ghi đè để tính hình tròn
        {
            return _height*_width;
        }
    }
}
