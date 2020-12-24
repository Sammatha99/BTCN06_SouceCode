using BTCN06_SouceCode;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace MyTest
{
    [TestClass]
    public class BaseClassTest
    {
        protected static string FileExcelInputTC = "TestCases.xlsx";
        protected static string SheetInputPointTC = "Sheet1";
        protected static string SheetInputTriangleTC = "Sheet2";
        protected static string SheetInputPointDistanceTC = "Sheet3";
        protected static string SheetInputTriangelIsTrueTC = "Sheet4";
        protected static string SheetInputTriangelIsAllPointsTrueTC = "Sheet5";

        protected static double Delta = 0.00001;

        #region Support_Methods
        public static IEnumerable<object[]> ReadPointExcelTC()
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(FileExcelInputTC)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[SheetInputPointTC];
                int rowCount = worksheet.Dimension.End.Row;
                for (int i = 2; i <= rowCount; i++)
                {
                    yield return new object[]
                    {
                        worksheet.Cells[i,1].Value?.ToString().Trim(),
                        worksheet.Cells[i,2].Value?.ToString().Trim(),
                        worksheet.Cells[i,3].Value?.ToString().Trim()
                    };
                }
            }
        }

        public static IEnumerable<object[]> ReadPointDistanceExcelTC()
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(FileExcelInputTC)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[SheetInputPointDistanceTC];
                int rowCount = worksheet.Dimension.End.Row;
                for (int i = 2; i <= rowCount; i++)
                {
                    yield return new object[]
                    {
                        worksheet.Cells[i,1].Value?.ToString().Trim(),
                        worksheet.Cells[i,2].Value?.ToString().Trim(),
                        worksheet.Cells[i,3].Value?.ToString().Trim(),
                        worksheet.Cells[i,4].Value?.ToString().Trim(),
                        worksheet.Cells[i,5].Value?.ToString().Trim()
                    };
                }
            }
        }

        public static IEnumerable<object[]> ReadTriangleExcelTC()
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(FileExcelInputTC)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[SheetInputTriangleTC];
                int rowCount = worksheet.Dimension.End.Row;
                for (int i = 2; i <= rowCount; i++)
                {
                    yield return new object[]
                    {
                        worksheet.Cells[i,1].Value?.ToString().Trim(),
                        worksheet.Cells[i,2].Value?.ToString().Trim(),
                        worksheet.Cells[i,3].Value?.ToString().Trim(),
                        worksheet.Cells[i,4].Value?.ToString().Trim(),
                        worksheet.Cells[i,5].Value?.ToString().Trim(),
                        worksheet.Cells[i,6].Value?.ToString().Trim(),
                        worksheet.Cells[i,7].Value?.ToString().Trim(),
                        worksheet.Cells[i,8].Value?.ToString().Trim()
                    };
                }
            }
        }

        public static IEnumerable<object[]> ReadIsaTriangleExcelTC()
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(FileExcelInputTC)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[SheetInputTriangelIsTrueTC];
                int rowCount = worksheet.Dimension.End.Row;
                for (int i = 2; i <= rowCount; i++)
                {
                    yield return new object[]
                    {
                        worksheet.Cells[i,1].Value?.ToString().Trim(),
                        worksheet.Cells[i,2].Value?.ToString().Trim(),
                        worksheet.Cells[i,3].Value?.ToString().Trim(),
                        worksheet.Cells[i,4].Value?.ToString().Trim(),
                        worksheet.Cells[i,5].Value?.ToString().Trim(),
                        worksheet.Cells[i,6].Value?.ToString().Trim(),
                        worksheet.Cells[i,7].Value?.ToString().Trim()
                    };
                }
            }
        }

        public static IEnumerable<object[]> ReadIsaTriangleAllPointIsTrueExcelTC()
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(FileExcelInputTC)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[SheetInputTriangelIsAllPointsTrueTC];
                int rowCount = worksheet.Dimension.End.Row;
                for (int i = 2; i <= rowCount; i++)
                {
                    yield return new object[]
                    {
                        worksheet.Cells[i,1].Value?.ToString().Trim(),
                        worksheet.Cells[i,2].Value?.ToString().Trim(),
                        worksheet.Cells[i,3].Value?.ToString().Trim(),
                        worksheet.Cells[i,4].Value?.ToString().Trim(),
                        worksheet.Cells[i,5].Value?.ToString().Trim(),
                        worksheet.Cells[i,6].Value?.ToString().Trim(),
                        worksheet.Cells[i,7].Value?.ToString().Trim()
                    };
                }
            }
        }
        #endregion

        #region PointTest
        Point point1;
        Point point2;



        [TestMethod]
        public void IsAPointTest_Test()
        {
            Assert.AreEqual(true, new Point("1", "a").IsAPoint());
        }

        [TestMethod]
        public void Fail_Test()
        {
            Assert.Fail();
        }

        [DynamicData(nameof(ReadPointExcelTC), DynamicDataSourceType.Method)]
        [TestMethod]
        [DataTestMethod]
        public void IsAPointTest(string x, string y, string result)
        {
            bool IsPoint = result == "0" ? false : true;
            point1 = new Point(x, y);
            Assert.AreEqual(IsPoint, point1.IsAPoint());
        }

        [DynamicData(nameof(ReadPointDistanceExcelTC), DynamicDataSourceType.Method)]
        [TestMethod]
        [DataTestMethod]
        public void GetDistanceTest(string x1, string y1, string x2, string y2, string expectedResult)
        {
            point1 = new Point(x1, y1);
            point2 = new Point(x2, y2);
            double result = Point.getDistanceBetween2Points(point1, point2);
            double ExpectedResult = double.Parse(expectedResult, CultureInfo.InvariantCulture.NumberFormat);
            Assert.IsTrue(Math.Abs(result - ExpectedResult) < Delta);
        }
        #endregion

        #region TriangleTest
        Triangle triangle;

        [DynamicData(nameof(ReadIsaTriangleExcelTC), DynamicDataSourceType.Method)]
        [TestMethod]
        [DataTestMethod]
        public void IsATriangleTest(string x1, string y1, string x2, string y2, string x3, string y3, string result)
        {
            triangle = new Triangle( new List<Point> {
                new Point(x1,y1),
                new Point(x2,y2),
                new Point(x3,y3)});
            bool IsTriangle = result == "0" ? false : true;
            triangle.edge1 = Point.getDistanceBetween2Points(new Point(x1 , y1), new Point(x2, y2));
            triangle.edge2 = Point.getDistanceBetween2Points(new Point(x1 , y1), new Point(x3, y3));
            triangle.edge3 = Point.getDistanceBetween2Points(new Point(x2 , y2), new Point(x3, y3));
            Assert.AreEqual(IsTriangle, triangle.IsATriangle());
        }

        [DynamicData(nameof(ReadTriangleExcelTC), DynamicDataSourceType.Method)]
        [TestMethod]
        [DataTestMethod]
        public void TestTriangleType(string x1, string y1, string x2, string y2, string x3, string y3, string ExpectedType, string ExpectedPer)
        {
            triangle = new Triangle(new List<Point> {
                new Point(x1,y1),
                new Point(x2,y2),
                new Point(x3,y3)});
            triangle.edge1 = Point.getDistanceBetween2Points(new Point(x1, y1), new Point(x2, y2));
            triangle.edge2 = Point.getDistanceBetween2Points(new Point(x1, y1), new Point(x3, y3));
            triangle.edge3 = Point.getDistanceBetween2Points(new Point(x2, y2), new Point(x3, y3));
            Assert.AreEqual(ExpectedType.Trim().ToLower(), triangle.GetTypeOfTriangle().Trim().ToLower());
        }

        [DynamicData(nameof(ReadTriangleExcelTC), DynamicDataSourceType.Method)]
        [TestMethod]
        [DataTestMethod]
        public void TestTrianglePer(string x1, string y1, string x2, string y2, string x3, string y3, string ExpectedType, string ExpectedPer)
        {
            triangle = new Triangle(new List<Point> {
                new Point(x1,y1),
                new Point(x2,y2),
                new Point(x3,y3)});
            triangle.edge1 = Point.getDistanceBetween2Points(new Point(x1, y1), new Point(x2, y2));
            triangle.edge2 = Point.getDistanceBetween2Points(new Point(x1, y1), new Point(x3, y3));
            triangle.edge3 = Point.getDistanceBetween2Points(new Point(x2, y2), new Point(x3, y3));
            double newExpectedPer = double.Parse(ExpectedPer, CultureInfo.InvariantCulture.NumberFormat);
            Assert.IsTrue(Math.Abs(newExpectedPer - triangle.GetPerimeter()) < Delta);
        }

        [DynamicData(nameof(ReadIsaTriangleAllPointIsTrueExcelTC), DynamicDataSourceType.Method)]
        [TestMethod]
        [DataTestMethod]
        public void TestTriangleIsAllPoint(string x1, string y1, string x2, string y2, string x3, string y3, string ExpectedResult)
        {
            triangle = new Triangle(new List<Point> {
                new Point(x1,y1),
                new Point(x2,y2),
                new Point(x3,y3)});
            Assert.AreEqual(ExpectedResult.ToLower(), triangle.IsAllPoint().Trim().ToLower());
        }
        #endregion
    }
}
