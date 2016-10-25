// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at http://mozilla.org/MPL/2.0/.

using ExcelDna.Integration;
using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace Xl {

    public struct XlResult<TResult> : IDisposable {
        private Func<TResult> _func;
        private TaskCompletionSource<TResult> _tcs;

        public XlResult(Func<TResult> func) {
            _func = func;
            _tcs = null;
        }

        public XlResult(TResult value)
        {
            // TODO interesting if this pattern avoids Func allocation
            // the struct should be stored somewhere, is CLR smart enough
            // to keep it without boxing?
            var identity = new IdentityResult<TResult>(value);
            _func = identity.Call; // was () => value; this obviously captures value and will allocate
            _tcs = new TaskCompletionSource<TResult>();
            _tcs.SetResult(value);
        }

        // awaiting on XlResult is thread-safe and only call QueueAsMacro once
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public XlResult<TCont> Map<TCont>(Func<XlResult<TResult>, TCont> continuation) {
            // When this f is called (e.g. via await) on non-main thread,
            // f is queued as macro. Then GetResult() of this instance is
            // called and since we are now on the main thread, it will return
            // instantly, and then we could access its result many times
            // and no intermediate COM objects are created
            var th = this;
            Func<TCont> f = () => {
                th.GetResult(true);
                return continuation(th);
            };
            return new XlResult<TCont>(f);
        }

        // simple function composition, evaluation is delayed until a result is requested
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public XlResult<TCont> Map<TCont>(Func<TResult, TCont> continuation) {
            var th = this;
            Func<TCont> f = () => {
                // this line will return synchronously
                var result = th.GetResult(true).Result;
                return continuation(result);
            };
            return new XlResult<TCont>(f);
        }

        public bool IsCompleted
        {
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            get
            {
                var tcs = _tcs;
                if (tcs == null) return false;
                return tcs.Task.IsCompleted;
            }
        }

        public TResult Result
        {
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            get { return GetResult(false).Result; }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private Task<TResult> GetResult(bool ensureMainThread) {
            var alreadyStarted = false;
            var func = Volatile.Read(ref _func);
            if (func == null) {
                alreadyStarted = true;
            } else {
                lock (func) {
                    if (_tcs == null) {
                        _tcs = new TaskCompletionSource<TResult>();
                    } else {
                        alreadyStarted = true;
                    }
                }
            }
            if (alreadyStarted) return _tcs.Task;

            if (Thread.CurrentThread.ManagedThreadId == 1) {
                var result = func();
                _tcs.TrySetResult(result);
            } else {
                if (ensureMainThread) throw new InvalidOperationException("Expected to already run on main thread");
                var th = this;
                ExcelAsyncUtil.QueueAsMacro(() => {
                    var result = func();
                    th._tcs.TrySetResult(result);
                });
            }
            // release the reference to let GC collect it and anything it has captured
            Volatile.Write(ref _func, null);
            return _tcs.Task;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public TaskAwaiter<TResult> GetAwaiter() {
            return GetResult(false).GetAwaiter();
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public ConfiguredTaskAwaitable<TResult> ConfigureAwait(bool continueOnCapturedContext) {
            return GetResult(false).ConfigureAwait(continueOnCapturedContext);
        }

        public void Dispose() {
            Dispose(true);
        }

        public void Dispose(bool disposing) {
            var tcs = _tcs;
            if (tcs == null || !tcs.Task.IsCompleted) return;
            var result = _tcs.Task.Result;
            if (result != null && Marshal.IsComObject(result)) {
                Marshal.ReleaseComObject(result);
            }
            _tcs = null;
        }
    }


    internal struct IdentityResult<T> {
        public T Value;

        public IdentityResult(T value) {
            Value = value;
        }

        public T Call() {
            return Value;
        }
    }

    public struct XlCallResult {
        private readonly int _function;
        private object[] _parameters;

        public XlCall.XlReturn XlReturn { get; private set; }

        public object Result { get; private set; }

        public XlCallResult(int function, params object[] parameters) {
            _function = function;
            _parameters = parameters;
            Result = null;
            XlReturn = default(XlCall.XlReturn);
        }

        internal object Call() {
            try {
                return XlCall.Excel(_function, _parameters);
            } finally {
                _parameters = null;
            }
        }

        internal XlCallResult TryCall() {
            try {
                object result;
                XlReturn = XlCall.TryExcel(_function, out result, _parameters);
                Result = result;
                return this;
            } finally {
                _parameters = null;
            }
        }
    }

    public class XlResult {

        /// <summary>
        /// Call ExcelDna's XlCall.Excel function asynchronously on the main thread
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static XlResult<object> XlCall(int function, params object[] parameters) {
            var xlCallParameters = new XlCallResult(function, parameters);
            return new XlResult<object>(xlCallParameters.Call);
        }

        /// <summary>
        /// Call ExcelDna's XlCall.TryExcel function asynchronously on the main thread
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static XlResult<XlCallResult> XlTryCall(int function, params object[] parameters) {
            var xlCallParameters = new XlCallResult(function, parameters);
            return new XlResult<XlCallResult>(xlCallParameters.TryCall);
        }
    }
}