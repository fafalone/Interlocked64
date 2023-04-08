# Interlocked64
## Interlocked64.dll v1.0.1
x64 Interlocked* compiler instrinsics exported by standard DLL

The Interlocked functions are very useful for multithreading, but are not exported under x64. I suspect a toolchain lockin strategy. As these functions are not trivially recreated using higher level code than assembly (when you run down those functions, they hit intrlock.asm), I've created a simple C++ DLL to turn them back into exported functions. This is primarily intended for use in twinBASIC, but if anyone did come up with a use for it in VBA7 64bit, it would work there too-- it's just a standard DLL.

The demo application includes all of the declares, and a very simple example of using a couple. 

**NOTE:** The Demo, as configured, only works when compiled:

VBx and tB don't search the app path for DLLs by default.

The typical workaround is to hard code the path, but who knows where you'll put this? An alternative is provided here **BUT IT ONLY WORKS FOR COMPILED EXES**.
To use these functions from the IDE, or without the Set dir call, you can either hard code the path or place the DLL in System32.

### Declarations
Here's the full set of declares for the DLL:

```
    Private DeclareWide PtrSafe Function InterlockedIncrement Lib "Interlocked64.dll" Alias "x64InterlockedIncrement" (Addend As Long) As Long
    Private DeclareWide PtrSafe Function InterlockedDecrement Lib "Interlocked64.dll" Alias "x64InterlockedDecrement" (Addend As Long) As Long
    Private DeclareWide PtrSafe Function InterlockedIncrement16 Lib "Interlocked64.dll" Alias "x64InterlockedIncrement16" (Addend As Integer) As Integer
    Private DeclareWide PtrSafe Function InterlockedDecrement16 Lib "Interlocked64.dll" Alias "x64InterlockedDecrement16" (Addend As Integer) As Integer
    Private DeclareWide PtrSafe Function InterlockedIncrement64 Lib "Interlocked64.dll" Alias "x64InterlockedIncrement64" (Addend As LongLong) As LongLong
    Private DeclareWide PtrSafe Function InterlockedDecrement64 Lib "Interlocked64.dll" Alias "x64InterlockedDecrement64" (Addend As LongLong) As LongLong
    
    Private DeclareWide PtrSafe Function InterlockedExchange Lib "Interlocked64.dll" Alias "x64InterlockedExchange" (target As Long, ByVal value As Long) As Long
    Private DeclareWide PtrSafe Function InterlockedExchangeRef Lib "Interlocked64.dll" Alias "x64InterlockedExchange64" (target As LongLong, value As Any) As LongLong
    Private DeclareWide PtrSafe Function InterlockedExchange8 Lib "Interlocked64.dll" Alias "x64InterlockedExchange8" (target As Byte, ByVal value As Byte) As Byte
    Private DeclareWide PtrSafe Function InterlockedExchange16 Lib "Interlocked64.dll" Alias "x64InterlockedExchange16" (Destination As Integer, ByVal ExChange As Integer) As Integer
    Private DeclareWide PtrSafe Function InterlockedExchange64 Lib "Interlocked64.dll" Alias "x64InterlockedExchange64" (target As LongLong, ByVal value As LongLong) As LongLong
    
    Private DeclareWide PtrSafe Function InterlockedExchangePointer Lib "Interlocked64.dll" Alias "x64InterlockedExchangePointer" (target As Any, value As Any) As LongPtr
    
    Private DeclareWide PtrSafe Function InterlockedExchangeAdd Lib "Interlocked64.dll" Alias "x64InterlockedExchangeAdd" (Addend As Long, ByVal value As Long) As Long
    Private DeclareWide PtrSafe Function InterlockedExchangeAdd64 Lib "Interlocked64.dll" Alias "x64InterlockedExchangeAdd64" (Addend As LongLong, ByVal value As LongLong) As LongLong
   
    Private DeclareWide PtrSafe Function InterlockedAdd Lib "Interlocked64.dll" Alias "x64InterlockedAdd" (Addend As Long, ByVal value As Long) As Long
    Private DeclareWide PtrSafe Function InterlockedAdd64 Lib "Interlocked64.dll" Alias "x64InterlockedAdd64" (Addend As LongLong, ByVal value As LongLong) As LongLong
    
    Private DeclareWide PtrSafe Function InterlockedAnd Lib "Interlocked64.dll" Alias "x64InterlockedAnd" (Destination As Long, ByVal value As Long) As Long
    Private DeclareWide PtrSafe Function InterlockedAnd8 Lib "Interlocked64.dll" Alias "x64InterlockedAnd8" (Destination As Byte, ByVal value As Byte) As Byte
    Private DeclareWide PtrSafe Function InterlockedAnd16 Lib "Interlocked64.dll" Alias "x64InterlockedAnd16" (Destination As Integer, ByVal value As Integer) As Integer
    Private DeclareWide PtrSafe Function InterlockedAnd64 Lib "Interlocked64.dll" Alias "x64InterlockedAnd64" (Destination As LongLong, ByVal value As LongLong) As LongLong
    
    Private DeclareWide PtrSafe Function InterlockedOr Lib "Interlocked64.dll" Alias "x64InterlockedOr" (Destination As Long, ByVal value As Long) As Long
    Private DeclareWide PtrSafe Function InterlockedOr8 Lib "Interlocked64.dll" Alias "x64InterlockedOr8" (Destination As Byte, ByVal value As Byte) As Byte
    Private DeclareWide PtrSafe Function InterlockedOr16 Lib "Interlocked64.dll" Alias "x64InterlockedOr16" (Destination As Integer, ByVal value As Integer) As Integer
    Private DeclareWide PtrSafe Function InterlockedOr64 Lib "Interlocked64.dll" Alias "x64InterlockedOr64" (Destination As LongLong, ByVal value As LongLong) As LongLong

    Private DeclareWide PtrSafe Function InterlockedXOr Lib "Interlocked64.dll" Alias "x64InterlockedXOr" (Destination As Long, ByVal value As Long) As Long
    Private DeclareWide PtrSafe Function InterlockedXOr8 Lib "Interlocked64.dll" Alias "x64InterlockedXOr8" (Destination As Byte, ByVal value As Byte) As Byte
    Private DeclareWide PtrSafe Function InterlockedXOr16 Lib "Interlocked64.dll" Alias "x64InterlockedXOr16" (Destination As Integer, ByVal value As Integer) As Integer
    Private DeclareWide PtrSafe Function InterlockedXOr64 Lib "Interlocked64.dll" Alias "x64InterlockedXOr64" (Destination As LongLong, ByVal value As LongLong) As LongLong

    Private DeclareWide PtrSafe Function InterlockedCompareExchange Lib "Interlocked64.dll" Alias "x64InterlockedCompareExchange" (Destination As Long, ByVal Exchange As Long, ByVal Comperand As Long) As Long
    Private DeclareWide PtrSafe Function InterlockedCompareExchange16 Lib "Interlocked64.dll" Alias "x64InterlockedCompareExchange16" (Destination As Integer, ByVal Exchange As Integer, ByVal Comperand As Integer) As Integer
    Private DeclareWide PtrSafe Function InterlockedCompareExchange64 Lib "Interlocked64.dll" Alias "x64InterlockedCompareExchange64" (Destination As LongLong, ByVal Exchange As LongLong, ByVal Comperand As LongLong) As LongLong
    Private DeclareWide PtrSafe Function InterlockedCompareExchange128 Lib "Interlocked128.dll" Alias "x64InterlockedCompareExchange128" (Destination As LongLong, ByVal ExchangeHigh As LongLong, ByVal ExchangeLow As LongLong, ByVal Comperand As Integer) As Byte
    Private DeclareWide PtrSafe Function InterlockedCompareExchangePointer Lib "Interlocked64.dll" Alias "x64InterlockedCompareExchangePointer" (Destination As Any, Exchange As Any, Comperand As Any) As LongPtr
```
