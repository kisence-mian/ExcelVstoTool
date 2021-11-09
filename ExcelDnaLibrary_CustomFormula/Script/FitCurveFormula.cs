using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDnaLibrary_CustomFormula
{
    /// <summary>
    /// 拟合曲线，给定一个曲线公式，并给出起点和终点，自动填充中间的值
    /// </summary>
    public class FitCurveFormula
    {
        [ExcelFunction(Name = "FitCurve", Description = "给定一个曲线公式，并给出起点和终点，自动填充中间的值", Category = "拟合曲线")]
        public static double FitCurve(
             [ExcelArgument(Name = "Start", Description = "目标值")]double startValue,
             [ExcelArgument(Name = "End", Description = "默认值")]double endValue,
            [ExcelArgument(Name = "Current", Description = "当前步")]int current,
            [ExcelArgument(Name = "Step", Description = "总步长")]int step,
            [ExcelArgument(Name = "InterpType", Description = "插值算法")]string interpType)
        {
            double ct = current;
            double tt = step;
            InterpType it = (InterpType)Enum.Parse(typeof(InterpType), interpType);

            return GetInterpolation(it, startValue, endValue, ct, tt);
        }

        [ExcelFunction(Name = "FitCurve_AllInterpName", Description = "获取所有插值算法名称", Category = "拟合曲线")]
        public static string GetAllInterpName()
        {
            string result = "";

            foreach (InterpType item in Enum.GetValues(typeof(InterpType)))
            {
                result += item.ToString() + "\n";
            }

            return result;
        }

        #region 插值算法

        #region 总入口

        static double GetInterpolation(InterpType interpolationType, double oldValue, double aimValue, double currentTime, double totalTime)
        {
            switch (interpolationType)
            {
                case InterpType.Default:
                case InterpType.Linear: return Liner(oldValue, aimValue, currentTime, totalTime);

                case InterpType.InBack: return InBack(oldValue, aimValue, currentTime, totalTime);
                case InterpType.OutBack: return OutBack(oldValue, aimValue, currentTime, totalTime);
                case InterpType.InOutBack: return InOutBack(oldValue, aimValue, currentTime, totalTime);
                case InterpType.OutInBack: return OutInBack(oldValue, aimValue, currentTime, totalTime);

                case InterpType.InQuad: return InQuad(oldValue, aimValue, currentTime, totalTime);
                case InterpType.OutQuad: return OutQuad(oldValue, aimValue, currentTime, totalTime);
                case InterpType.InoutQuad: return InoutQuad(oldValue, aimValue, currentTime, totalTime);

                case InterpType.InCubic: return InCubic(oldValue, aimValue, currentTime, totalTime);
                case InterpType.OutCubic: return OutCubic(oldValue, aimValue, currentTime, totalTime);
                case InterpType.InoutCubic: return InoutCubic(oldValue, aimValue, currentTime, totalTime);
                case InterpType.OutInCubic: return OutinCubic(oldValue, aimValue, currentTime, totalTime);

                case InterpType.InQuart: return InQuart(oldValue, aimValue, currentTime, totalTime);
                case InterpType.OutQuart: return OutQuart(oldValue, aimValue, currentTime, totalTime);
                case InterpType.InOutQuart: return InOutQuart(oldValue, aimValue, currentTime, totalTime);
                case InterpType.OutInQuart: return OutInQuart(oldValue, aimValue, currentTime, totalTime);

                case InterpType.InQuint: return InQuint(oldValue, aimValue, currentTime, totalTime);
                case InterpType.OutQuint: return OutQuint(oldValue, aimValue, currentTime, totalTime);
                case InterpType.InOutQuint: return InOutQuint(oldValue, aimValue, currentTime, totalTime);
                case InterpType.OutInQuint: return OutInQuint(oldValue, aimValue, currentTime, totalTime);

                case InterpType.InSine: return InSine(oldValue, aimValue, currentTime, totalTime);
                case InterpType.OutSine: return OutSine(oldValue, aimValue, currentTime, totalTime);
                case InterpType.InOutSine: return InOutSine(oldValue, aimValue, currentTime, totalTime);
                case InterpType.OutInSine: return OutInSine(oldValue, aimValue, currentTime, totalTime);

                case InterpType.InExpo: return InExpo(oldValue, aimValue, currentTime, totalTime);
                case InterpType.OutExpo: return OutExpo(oldValue, aimValue, currentTime, totalTime);
                case InterpType.InOutExpo: return InOutExpo(oldValue, aimValue, currentTime, totalTime);
                case InterpType.OutInExpo: return OutInExpo(oldValue, aimValue, currentTime, totalTime);

                case InterpType.InBounce: return InBounce(oldValue, aimValue, currentTime, totalTime);
                case InterpType.OutBounce: return OutBounce(oldValue, aimValue, currentTime, totalTime);
                case InterpType.InOutBounce: return InOutBounce(oldValue, aimValue, currentTime, totalTime);
                case InterpType.OutInBounce: return OutInBounce(oldValue, aimValue, currentTime, totalTime);
            }

            return 0;
        }

        #endregion

        public static double Liner(double b, double to, double t, double d)
        {
            double c = to - b;
            t = t / d;

            return b + c * t;
        }

        public static double InBack(double b, double to, double t, double d, double s = 1.70158f)
        {
            double c = to - b;
            t = t / d;

            return c * t * t * ((s + 1) * t - s) + b;
        }

        public static double OutBack(double b, double to, double t, double d, double s = 1.70158f)
        {
            double c = to - b;

            t = t / d - 1;
            //Debug.LogWarning(c * (t * t * ((s + 1) * t + s) + 1) + b);
            return c * (t * t * ((s + 1) * t + s) + 1) + b;

        }

        public static double InOutBack(double b, double to, double t, double d, double s = 1.70158f)
        {
            double c = to - b;
            s = s * 1.525f;
            t = t / d * 2;
            if (t < 1)
                return c / 2 * (t * t * ((s + 1) * t - s)) + b;
            else
            {
                t = t - 2;
                return c / 2 * (t * t * ((s + 1) * t + s) + 2) + b;
            }
        }

        public static double OutInBack(double b, double to, double t, double d, double s = 1.70158f)
        {
            double c = to - b;
            if (t < d / 2)
            {
                return OutBack(t * 2, c / 2, d, s);
            }

            else
            {
                return InBack(t * 2 - d, b + c / 2, c / 2, d, s);
            }

        }

        public static double InQuad(double b, double to, double t, double d)
        {
            double c = to - b;
            t = t / d;
            return (double)(c * Math.Pow(t, 2) + b);
        }

        public static double OutQuad(double b, double to, double t, double d)
        {
            double c = to - b;
            t = t / d;
            return (double)(-c * t * (t - 2) + b);
        }

        public static double InoutQuad(double b, double to, double t, double d)
        {
            double c = to - b;
            t = t / d * 2;
            if (t < 1)
                return (double)(c / 2 * Math.Pow(t, 2) + b);
            else
                return -c / 2 * ((t - 1) * (t - 3) - 1) + b;

        }
        public static double InCubic(double b, double to, double t, double d)
        {
            double c = to - b;
            t = t / d;
            return (double)(c * Math.Pow(t, 3) + b);

        }
        public static double OutCubic(double b, double to, double t, double d)
        {
            double c = to - b;
            t = t / d - 1;
            return (double)(c * (Math.Pow(t, 3) + 1) + b);

        }
        public static double InoutCubic(double b, double to, double t, double d)
        {
            double c = to - b;
            t = t / d * 2;
            if (t < 1)
                return c / 2 * t * t * t + b;
            else
            {
                t = t - 2;
                return c / 2 * (t * t * t + 2) + b;
            }
        }

        public static double OutinCubic(double b, double to, double t, double d)
        {
            double c = to - b;

            if (t < d / 2)
            {
                return OutCubic(b, b + c / 2, t * 2, d);
            }
            else
            {
                return InCubic(b + c / 2, to, (t * 2) - d, d);
            }
        }

        public static double InQuart(double b, double to, double t, double d)
        {
            double c = to - b;
            t = t / d;
            return (double)(c * Math.Pow(t, 4) + b);

        }
        public static double OutQuart(double b, double to, double t, double d)
        {
            double c = to - b;
            t = t / d - 1;
            return (double)(-c * (Math.Pow(t, 4) - 1) + b);

        }

        public static double InOutQuart(double b, double to, double t, double d)
        {
            double c = to - b;
            t = t / d * 2;
            if (t < 1)
                return (double)(c / 2 * Math.Pow(t, 4) + b);
            else
            {
                t = t - 2;
                return (double)(-c / 2 * (Math.Pow(t, 4) - 2) + b);
            }

        }
        public static double OutInQuart(double b, double to, double t, double d)
        {
            if (t < d / 2)
            {
                double c = to - b;
                t *= 2;
                c *= 0.5f;
                t = t / d - 1;

                return (double)(-c * (Math.Pow(t, 4) - 1) + b);
            }
            else
            {
                double c = to - b;
                t = t * 2 - d;
                b = b + c * 0.5f;
                c *= 0.5f;
                t = t / d;


                return (double)(c * Math.Pow(t, 4) + b);

            }
        }

        public static double InQuint(double b, double to, double t, double d)
        {
            double c = to - b;
            t = t / d;
            return (double)(c * Math.Pow(t, 5) + b);

        }

        public static double OutQuint(double b, double to, double t, double d)
        {
            double c = to - b;
            t = t / d - 1;
            return (double)(c * (Math.Pow(t, 5) + 1) + b);
        }

        public static double InOutQuint(double b, double to, double t, double d)
        {
            double c = to - b;
            t = t / d * 2;
            if (t < 1)
                return (double)(c / 2 * Math.Pow(t, 5) + b);
            else
            {
                t = t - 2;
                return (double)(c / 2 * (Math.Pow(t, 5) + 2) + b);

            }

        }

        public static double OutInQuint(double b, double to, double t, double d)
        {
            double c = to - b;
            if (t < d / 2)
            {
                t *= 2;
                c *= 0.5f;
                t = t / d - 1;
                return (double)(c * (Math.Pow(t, 5) + 1) + b);
            }
            else
            {
                t = t * 2 - d;
                b = b + c * 0.5f;
                c *= 0.5f;

                t = t / d;
                return (double)(c * Math.Pow(t, 5) + b);
            }
        }

        public static double InSine(double b, double to, double t, double d)
        {
            double c = to - b;
            return (double)(-c * Math.Cos(t / d * (Math.PI / 2)) + c + b);

        }

        public static double OutSine(double b, double to, double t, double d)
        {
            double c = to - b;
            return (double)(c * Math.Sin(t / d * (Math.PI / 2)) + b);
        }

        public static double InOutSine(double b, double to, double t, double d)
        {
            double c = to - b;
            return (double)(-c / 2 * (Math.Cos(Math.PI * t / d) - 1) + b);

        }
        public static double OutInSine(double b, double to, double t, double d)
        {
            double c = to - b;
            if (t < d / 2)
            {
                t *= 2;
                c *= 0.5f;
                return (double)(c * Math.Sin(t / d * (Math.PI / 2)) + b);
            }
            else
            {
                t = t * 2 - d;
                b += c * 0.5f;
                c *= 0.5f;
                return (double)(-c * Math.Cos(t / d * (Math.PI / 2)) + c + b);

            }
        }
        public static double InExpo(double b, double to, double t, double d)
        {
            double c = to - b;
            if (t == 0)
                return b;
            else
                return (double)(c * Math.Pow(2, 10 * (t / d - 1)) + b - c * 0.001f);
        }
        public static double OutExpo(double b, double to, double t, double d)
        {
            double c = to - b;
            if (t == d)
                return b + c;
            else
                return (double)(c * 1.001 * (-Math.Pow(2, -10 * t / d) + 1) + b);

        }
        public static double InOutExpo(double b, double to, double t, double d)
        {
            double c = to - b;
            if (t == 0)
                return b;
            if (t == d)
                return (b + c);

            t = t / d * 2;

            if (t < 1)
                return (double)(c / 2 * Math.Pow(2, 10 * (t - 1)) + b - c * 0.0005f);
            else
            {
                t = t - 1;
                return (double)(c / 2 * 1.0005 * (-Math.Pow(2, -10 * t) + 2) + b);

            }
        }

        public static double OutInExpo(double b, double to, double t, double d)
        {
            double c = to - b;
            if (t < d / 2)
            {
                t *= 2;
                c *= 0.5f;
                if (t == d)
                    return b + c;
                else
                    return (double)(c * 1.001 * (-Math.Pow(2, -10 * t / d) + 1) + b);
            }
            else
            {
                t = t * 2 - d;
                b += c * 0.5f;
                c *= 0.5f;
                if (t == 0)
                    return b;
                else
                    return (double)(c * Math.Pow(2, 10 * (t / d - 1)) + b - c * 0.001f);

            }
        }

        public static double OutBounce(double b, double to, double t, double d)
        {
            double c = to - b;
            t = t / d;
            if (t < 1 / 2.75)
            {
                return c * (7.5625f * t * t) + b;
            }
            else if (t < 2 / 2.75)
            {
                t = t - (1.5f / 2.75f);

                return c * (7.5625f * t * t + 0.75f) + b;
            }
            else if (t < 2.5 / 2.75)
            {

                t = t - (2.25f / 2.75f);
                return c * (7.5625f * t * t + 0.9375f) + b;
            }
            else
            {
                t = t - (2.625f / 2.75f);
                return c * (7.5625f * t * t + 0.984375f) + b;
            }
        }

        public static double InBounce(double b, double to, double t, double d)
        {
            double c = to - b;
            return c - OutBounce(0, to, d - t, d) + b;
        }

        public static double InOutBounce(double b, double to, double t, double d)
        {
            double c = to - b;
            if (t < d / 2)
            {
                return InBounce(0, to, t * 2f, d) * 0.5f + b;
            }
            else
            {
                return OutBounce(0, to, t * 2f - d, d) * 0.5f + c * 0.5f + b;
            }
        }

        public static double OutInBounce(double b, double to, double t, double d)
        {
            double c = to - b;
            if (t < d / 2)
            {
                return OutBounce(b, b + c / 2, t * 2, d);
            }
            else
            {
                return InBounce(b + c / 2, to, t * 2f - d, d);

            }
        }


        //outInExpo,
        //inBack,
        //outBack,
        //inOutBack,
        //outInBack,

        #endregion
    }

    //插值算法类型
    public enum InterpType
    {
        Default,
        Linear,
        InBack,
        OutBack,
        InOutBack,
        OutInBack,
        InQuad,
        OutQuad,
        InoutQuad,
        InCubic,
        OutCubic,
        InoutCubic,
        OutInCubic,
        InQuart,
        OutQuart,
        InOutQuart,
        OutInQuart,
        InQuint,
        OutQuint,
        InOutQuint,
        OutInQuint,
        InSine,
        OutSine,
        InOutSine,
        OutInSine,

        InExpo,
        OutExpo,
        InOutExpo,
        OutInExpo,

        OutBounce,
        InBounce,
        InOutBounce,
        OutInBounce,
    }
}
