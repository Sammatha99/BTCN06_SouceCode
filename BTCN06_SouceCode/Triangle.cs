using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BTCN06_SouceCode
{
    public class Triangle
    {
        List<Point> Points = new List<Point>();
        public double edge1 { get; set; }
        public double edge2 { get; set; }
        public double edge3 { get; set; }

        public static double delta = 0.0000001;

        public Triangle(List<Point> points)
        {
            Points = points;
        }


        public void Result()
        {
            if (!string.IsNullOrEmpty(IsAllPoint())) {
                Console.WriteLine(IsAllPoint());
                return;
            }
            edge1 = Point.getDistanceBetween2Points(Points[0], Points[1]);
            edge2 = Point.getDistanceBetween2Points(Points[0], Points[2]);
            edge3 = Point.getDistanceBetween2Points(Points[2], Points[1]);
            if (IsATriangle())
            {
                Console.WriteLine($"Loại tam giác:{GetTypeOfTriangle()}\n");
                Console.WriteLine($"Chu vi:{GetPerimeter()}\n");
                return;
            }
            Console.WriteLine("3 điểm trên không tạo thành 1 tam giác");
        }

        public string IsAllPoint()
        {
            for (int i = 0; i < 3; i++)
            {
                if (!Points[i].IsAPoint())
                {
                    return $"Điểm {i + 1} không hợp lệ";
                }
            }
            return "";
        }

        public bool IsATriangle()
        {
            return (edge1 < edge2 + edge3 &&
                    edge2 < edge1 + edge3 &&
                    edge3 < edge1 + edge2 );
        }

        public string GetTypeOfTriangle()
        {
            string result = "";
            if (CompareDouble(edge1, edge2, delta) &&
                CompareDouble(edge1, edge3, delta) &&
                CompareDouble( edge3,edge2,delta))
                result = result + "đều ";
            else
            {
                if (CompareDouble( Math.Pow(edge1, 2),(Math.Pow(edge2, 2) + Math.Pow(edge3, 2)), delta) ||
                    CompareDouble(Math.Pow(edge2, 2), (Math.Pow(edge1, 2) + Math.Pow(edge3, 2)),delta) ||
                    CompareDouble(Math.Pow(edge3, 2), (Math.Pow(edge2, 2) + Math.Pow(edge1, 2)),delta))
                    result = result + "vuông ";
                if (CompareDouble(edge1, edge2, delta) ||
                    CompareDouble(edge1, edge3, delta) ||
                    CompareDouble(edge3, edge2, delta))
                    result = result + "cân ";
            }

            return result != "" ? result : "thường";
        }

        public double GetPerimeter()
        {
            return edge1 + edge2 + edge3;
        }

        public static bool CompareDouble(double a, double b, double delta)
        {
            return Math.Abs(a - b) < delta;
        }
    }
}
