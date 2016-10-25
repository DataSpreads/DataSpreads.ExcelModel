## Excel Model

Rewriting my old Excel utils and code from [Joachim Loebb](https://github.com/smartquant/ExcelDna.Utilities)
using XlResult<T> pattern.
We wrap all functions and properties that require C/COM API call into XlResult<T>,
which works like Task but executes on the main thread instead of a thread pool.

* XlResult<T> is awaitable and works nicely inside async methods
* XlResult<T> is disposable and is a struct
* XlResult<T> Map() methods reuse cached results, mitigating COM object leakage (TODO tests)
* Result T is stored internally in a TaskCompletionSource.Task property. Since TCS is an
object, copying XlResult<T> will copy the reference to the TCS instance, while the reference 
to result T inside the Task remains unique. This simplifies managing COM objects.
* Currently XlResult<T> always allocates TCS, even if already on the main thread. Frequent 
call to C/COM API should be grouped together inside a single lambda. Intended usage of 
XlResult<T> is for heavy objects such as Workbooks, Ranges, etc.


## Sample

```
[ExcelFunction(IsMacroType = true)]
public static async Task<string> Hello([ExcelArgument(AllowReference = true)]object text) {
    var xlRef = text as ExcelReference;
    if (xlRef != null) {
        using (var result = XlResult.XlCall(XlCall.xlfReftext, xlRef, true)
            .Map(str => DSAddIn.XlApp.Evaluate((string)str) as Range)
            .Map(rng => rng.Value2.ToString())) {
            return await result;
        }
    }
    throw new Exception();
}
```

## Goal

The goal is to make Excel VBA/COM-like API that exposes only .NET types, preferably using
C API, but falling back on COM Interop internally and managing COM object without exposing them.
The API must be thread safe - both from concurrency and Excel-specific point of views:
* Concurrent access should not require locks, synchronization should be done internally.
* All C/COM API call are only made on the main Excel thread. Otherwise, C API throws almost
always, while COM API returns HRESULT busy result when a user edits a cell content and in other cases.

## Contributing

Contributions are welcome! At the beginning, I will only add functionality that I frequently use,
and XResult<T> usage could replace premature abstractions in most cases. If you want to add something, 
feel free to send a PR and add yourself to the authors list below.

This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0. If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

(c) Victor Baybekov, 2016