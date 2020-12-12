using System;
using ExcelDna.Integration;

namespace Excel_CEIDE2000
{
    /**
     * Implements functionality to calculate the CEIDE2000 color-difference and provides an interface for Excel.
     */
    public class CIEDE2000
    {
        private const string ExcelCategory = "Color Difference";

        private readonly double _l1S;
        private readonly double _a1S;
        private readonly double _b1S;

        private const double Pi = Math.PI;
        private const double Pi2 = 2.0 * Math.PI;

        private const double kL = 1.0, kC = 1.0, kH = 1.0;

        [ExcelFunction(Description = "Calculates the distance in the AB plane from a reference color (l1s, a1s, b1s) that has a specific DE00 difference and angle", Category = ExcelCategory)]
        public static double DE00ColorWithDifference(double l1s, double a1s, double b1s, double difference, double angle)
        {
            var colDist = new CIEDE2000(l1s, a1s, b1s);
            return colDist.ColorWithDifference(difference, angle);
        }

        [ExcelFunction(Description = "Calculates the A component of the color at a specific difference and angle from a reference color (l1, a1, b1).", Category = ExcelCategory)]
        public static double DE00A(double l1s, double a1s, double b1s, double difference, double angle)
        {
            var colDist = new CIEDE2000(l1s, a1s, b1s);
            var r = colDist.ColorWithDifference(difference, angle);
            return a1s + r * Math.Cos(angle);
        }

        [ExcelFunction(Description = "Calculates the B component of the color at a specific difference and angle from a reference color (l1, a1, b1).", Category = ExcelCategory)]
        public static double DE00B(double l1s, double a1s, double b1s, double difference, double angle)
        {
            var colDist = new CIEDE2000(l1s, a1s, b1s);
            var r = colDist.ColorWithDifference(difference, angle);
            return b1s + r * Math.Sin(angle);
        }

        [ExcelFunction(Description = "Calculates the DE2000 color-difference between two colors (l1, a1, b1) and (l2, a2, b2).", Category = ExcelCategory)]
        public static double DE00Difference(double l1s, double a1s, double b1s, double l2s, double a2s, double b2s)
        {
            var colDist = new CIEDE2000(l1s, a1s, b1s);
            return colDist.DE00(l2s, a2s, b2s);
        }


        [ExcelFunction(Description = "Calculates the DE2000 color-difference the color (l1, a1, b1) and the color that is \"radius\" away in direction \"angle\".", Category = ExcelCategory)]
        public static double DE00DifferencePolar(double l1s, double a1s, double b1s, double radius, double angle)
        {
            var colDist = new CIEDE2000(l1s, a1s, b1s);
            return colDist.DE00Polar(radius, angle);
        }

        /**
         * Constructor that takes the reference color (L, a, b) and all calculations are done
         * using this as reference color.
         */
        public CIEDE2000(double l1SIn, double a1SIn, double b1SIn)
        {
            _l1S = l1SIn;
            _a1S = a1SIn;
            _b1S = b1SIn;
        }

        public double ColorWithDifference(double difference, double angle)
        {
            angle = NormalizeAngle(angle);
            var f = new Func<double, double>((r) => DE00Polar(r, angle) - difference);

            var r1 = 0.0;
            var r2 = 2.0;
            // Try to find a large enough upper bound
            for(var i = 0; i < 10; i++)
            {
                if (f(r2) < 0)
                {
                    r2 *= 2.0;
                }
                else break;
            }
            // Simple second root finding. Stolen from www.geeksforgeeks.org
            var eps = 0.0001;
            double n = 0, xm, x0 = -1.0, c;
            if (f(r1) * f(r2) < 0)
            {
                do
                {

                    // calculate the intermediate 
                    // value 
                    x0 = (r1 * f(r2) - r2 * f(r1))
                        / (f(r2) - f(r1));

                    // check if x0 is root of 
                    // equation or not 
                    c = f(r1) * f(x0);

                    // update the value of interval 
                    r1 = r2;
                    r2 = x0;

                    // update number of iteration 
                    n++;

                    // if x0 is the root of equation  
                    // then break the loop 
                    if (c == 0)
                        break;
                    xm = (r1 * f(r2) - r2 * f(r1))
                        / (f(r2) - f(r1));

                    // repeat the loop until  
                    // the convergence  
                } while (Math.Abs(xm - x0) >= eps);
            }
            return x0;
        }

        public double DE00Polar(double radius, double angle)
        {
            var a2s = _a1S + radius * Math.Cos(angle);
            var b2s = _b1S + radius * Math.Sin(angle);
            return DE00(_l1S, a2s, b2s);
        }

        /**
         * Implementation of 
         * "The CIEDE2000 Color-Difference Formula: Implementation Notes, Supplementary Test Data, and Mathematical Observations".
         */
        public double DE00(double l2s, double a2s, double b2s)
        {
            var mCs = (Math.Sqrt(_a1S * _a1S + _b1S * _b1S) + Math.Sqrt(a2s * a2s + b2s * b2s)) / 2.0;
            var G = 0.5 * (1.0 - Math.Sqrt(Math.Pow(mCs, 7) / (Math.Pow(mCs, 7) + Math.Pow(25.0, 7))));
            var a1p = (1.0 + G) * _a1S;
            var a2p = (1.0 + G) * a2s;
            var C1p = Math.Sqrt(a1p * a1p + _b1S * _b1S);
            var C2p = Math.Sqrt(a2p * a2p + b2s * b2s);

            var h1p = Math.Abs(a1p) + Math.Abs(_b1S) > double.Epsilon ? Math.Atan2(_b1S, a1p) : 0.0;
            if (h1p < 0.0) h1p += Pi2;
            var h2p = Math.Abs(a2p) + Math.Abs(b2s) > double.Epsilon ? Math.Atan2(b2s, a2p) : 0.0;
            if (h2p < 0.0) h2p += Pi2;

            var dLp = l2s - _l1S;
            var dCp = C2p - C1p;

            var dhp = 0.0;
            var cProdAbs = Math.Abs(C1p * C2p);
            if (cProdAbs > double.Epsilon && Math.Abs(h1p - h2p) <= Pi)
            {
                dhp = h2p - h1p;
            }
            else if (cProdAbs > double.Epsilon && h2p - h1p > Pi)
            {
                dhp = h2p - h1p - Pi2;
            }
            else if (cProdAbs > Double.Epsilon && h2p - h1p < -Pi)
            {
                dhp = h2p - h1p + Pi2;
            }

            var dHp = 2.0 * Math.Sqrt(C1p * C2p) * Math.Sin(dhp / 2.0);

            var mLp = (_l1S + l2s) / 2.0;
            var mCp = (C1p + C2p) / 2.0;

            var mhp = 0.0;
            if (cProdAbs > double.Epsilon && Math.Abs(h1p - h2p) <= Pi)
            {
                mhp = (h1p + h2p) / 2.0;
            }
            else if (cProdAbs > double.Epsilon && Math.Abs(h1p - h2p) > Pi && h1p + h2p < Pi2)
            {
                mhp = (h1p + h2p + Pi2) / 2.0;
            }
            else if (cProdAbs > double.Epsilon && Math.Abs(h1p - h2p) > Pi && h1p + h2p >= Pi2)
            {
                mhp = (h1p + h2p - Pi2) / 2.0;
            }
            else if (cProdAbs <= double.Epsilon)
            {
                mhp = h1p + h2p;
            }

            var T = 1.0 - 0.17 * Math.Cos(mhp - Pi / 6.0) + .24 * Math.Cos(2.0 * mhp) +
                0.32 * Math.Cos(3.0 * mhp + Pi / 30.0) - 0.2 * Math.Cos(4.0 * mhp - 7.0 * Pi / 20.0);
            var dTheta = Pi / 6.0 * Math.Exp(-Math.Pow((mhp / (2.0 * Pi) * 360.0 - 275.0) / 25.0, 2));
            var RC = 2.0 * Math.Sqrt(Math.Pow(mCp, 7) / (Math.Pow(mCp, 7) + Math.Pow(25.0, 7)));
            var mlpSqr = (mLp - 50.0) * (mLp - 50.0);
            var SL = 1.0 + 0.015 * mlpSqr / Math.Sqrt(20.0 + mlpSqr);
            var SC = 1.0 + 0.045 * mCp;
            var SH = 1.0 + 0.015 * mCp * T;
            var RT = -Math.Sin(2.0 * dTheta) * RC;

            var de00 = Math.Sqrt(
                Math.Pow(dLp / (kL * SL), 2) + Math.Pow(dCp / (kC * SC), 2) + Math.Pow(dHp / (kH * SH), 2) +
                RT * dCp / (kC * SC) * dHp / (kH * SH)
            );
            return de00;
        }

        /**
         * Didn't find a better way to do this in C#.
         * It basically does what Mathematica does with Mod[angle, 2Pi].
         */
        private static double NormalizeAngle(double angle)
        {
            while (angle < 0.0)
            {
                angle += Pi2;
            }

            while (angle > Pi2)
            {
                angle -= Pi2;
            }

            return angle;
        }

    }
}
