using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SendTtsAndSmsSystem
{
    class AsyncTemplate
    {
        public static Action OnInvokeStarting { get; set; }

        public static Action OnInvokeEnding { get; set; }

        public static void DoWorkAsync(Action beginAction, Action endAction, Action<Exception> errorAction)
        {
            ThreadPool.QueueUserWorkItem(new WaitCallback(
                (o) =>
                {
                    try
                    {
                        beginAction();
                        endAction();
                    }
                    catch (Exception ex)
                    {
                        errorAction(ex);
                        return;
                    }
                    finally
                    {
                        if (OnInvokeEnding != null)
                        {
                            OnInvokeEnding();
                        }
                    }
                })
            , null);

            if (OnInvokeStarting != null)
            {
                OnInvokeStarting();
            }
        }

        public static void DoWorkAsync<TResult>(Func<TResult> beginAction, Action<TResult> endAction, Action<Exception> errorAction)
        {
            ThreadPool.QueueUserWorkItem(new WaitCallback(
                (o) =>
                {
                    TResult result = default(TResult);

                    try
                    {
                        result = beginAction();

                        endAction(result);
                    }
                    catch (Exception ex)
                    {
                        errorAction(ex);
                        return;
                    }
                    finally
                    {
                        if (OnInvokeEnding != null)
                        {
                            OnInvokeEnding();
                        }
                    }
                })
            , null);

            if (OnInvokeStarting != null)
            {
                OnInvokeStarting();
            }
        }
    }
}
