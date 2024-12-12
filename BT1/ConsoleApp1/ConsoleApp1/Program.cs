using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ConsoleApp1;
using Inventor;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(" nhập số thứ nhất: ");
            int num1 = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine(" nhập số thứ hai: ");
            int num2 = Convert.ToInt32(Console.ReadLine());

            int tong= sum(num1, num2);

            Program p = new Program(); //nếu ko có từ satic
            int tong2 = p.noStaticSum(num1, num2);
            
            
        }
        
        static int sum (int a, int b) // phương thức tĩnh Static
        {
            return a + b;
        }

        int noStaticSum(int a, int b)
        {
            return a + b;
        }

    }
}
