using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BTCN06_SouceCode
{
    class Program
    {
        static List<Point> points = new List<Point>();
        static Triangle triangle;
        static string result;
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.Unicode;
            Console.InputEncoding = Encoding.Unicode;

            while (true)
            {
                for (int i = 0; i < 3; i++)
                {
                    points.Add(getPointNumberX(i + 1));
                }
                Console.WriteLine("Thực hiện tính toán ?\n" +
                    "\t-0: ok\n" +
                    "\t-khác: nhập lại");
                result = Console.ReadLine();
                if (result == "0")
                {
                    //gọi hàm tính toán
                    triangle = new Triangle(points);
                    triangle.Result();

                    Console.WriteLine("Tiếp tục chương trình ?\n" +
                   "\t-0: ok\n" +
                   "\t-khác: tắt");
                    result = Console.ReadLine();
                    if (result != "0")
                        break;
                }
                CleanAll();
            }
        }

        public static Point getPointNumberX(int i)
        {
            Console.WriteLine($"Điểm số {i}, hoành độ: ");
            string x = Console.ReadLine();
            Console.WriteLine($"Điểm số {i}, tung độ: ");
            string y= Console.ReadLine();
            return new Point(x, y);
        }

        public static void CleanAll()
        {
            points = new List<Point>();
            triangle = null;
            Console.Clear();
        }
    }
}
