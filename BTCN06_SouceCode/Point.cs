using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BTCN06_SouceCode
{
    public class Point
    {
        public string X { get; set; }
        public string Y { get; set; }

        public Point(string x, string y)
        {
            X = x;
            Y = y;
        }

        public bool IsAPoint()
        {
            try
            {
                double.Parse(X, CultureInfo.InvariantCulture.NumberFormat);
                double.Parse(Y, CultureInfo.InvariantCulture.NumberFormat);
                return true;
            }
            catch(Exception ex)
            {
                return false;
            }
        }

        public static double getDistanceBetween2Points(Point a, Point b)
        {
            double ax = double.Parse(a.X, CultureInfo.InvariantCulture.NumberFormat);
            double ay = double.Parse(a.Y, CultureInfo.InvariantCulture.NumberFormat);
            double bx = double.Parse(b.X, CultureInfo.InvariantCulture.NumberFormat);
            double by = double.Parse(b.Y, CultureInfo.InvariantCulture.NumberFormat);

            return Math.Sqrt(Math.Pow(ax - bx, 2) + Math.Pow(ay - by, 2));
        }
    }
}
