namespace ExcelProfiler
{
    using System;
    using System.Diagnostics;

    public class MicroStopwatch : Stopwatch
    {
        readonly double _microSecPerTick = 1000000D / Stopwatch.Frequency;
        
        public MicroStopwatch()
        {
            if (!Stopwatch.IsHighResolution)
            {
                throw new Exception("On this system the high-resolution performance counter is not available");
            }
        }

        public long ElapsedMicroseconds
        {
            get
            {
                return (long)(ElapsedTicks * _microSecPerTick);
            }
        }

        public double ElapsedMillisecondsHighResolution
        {
            get
            {
                return Math.Round(this.ElapsedMicroseconds / 1000.0, 2);
            }
        }

        public static MicroStopwatch StartNewMicroStopwatch()
        {
            MicroStopwatch timer = new MicroStopwatch();
            timer.Start();
            return timer;
        }
    }
}
