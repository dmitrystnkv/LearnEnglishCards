using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace LearnEnglishCards
{
    public static class Actions
    {
        public static void RunAfter(this Action action, TimeSpan span)
        {
            var dispatcherTimer = new DispatcherTimer { Interval = span };
            dispatcherTimer.Tick += (sender, args) =>
            {
                var timer = sender as DispatcherTimer;
                if (timer != null)
                {
                    timer.Stop();
                }

                action();
            };
            dispatcherTimer.Start();
        }
    }

    //<Namespace>.Utilities
    public static class CommonUtil
    {
        public static void Run(System.Action action, TimeSpan afterSpan)
        {
            action.RunAfter(afterSpan);
        }
    }
}
