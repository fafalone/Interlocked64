/*
Interlocked64.dll - Access to Interlocked* functions for x64 app outside of the MS modern toolchain.
This simply wraps what are available only as compiler intrinsics for x64, but available as normal
on x86, which has become a problem for VB programmer of late since twinBASIC supports x64. 

(c) 2023 Jon Johnson (fafalone)
http://www.github.com/fafalone/Interlocked64
*/
#include "pch.h"

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}

long __declspec (dllexport) __stdcall x64InterlockedIncrement(long* Addend)
{
    return InterlockedIncrement(Addend);
};

short __declspec (dllexport) __stdcall x64InterlockedIncrement16(short* Addend)
{
    return InterlockedIncrement16(Addend);
};

LONG64 __declspec (dllexport) __stdcall x64InterlockedIncrement64(LONG64* Addend)
{
    return InterlockedIncrement64(Addend);
};

long __declspec (dllexport) __stdcall x64InterlockedDecrement(long* Addend)
{
    return InterlockedDecrement(Addend);
};

LONG64 __declspec (dllexport) __stdcall x64InterlockedDecrement64(LONG64* Addend)
{
    return InterlockedDecrement64(Addend);
};

short __declspec (dllexport) __stdcall x64InterlockedDecrement16(short* Addend)
{
    return InterlockedDecrement16(Addend);
};



long __declspec (dllexport) __stdcall x64InterlockedExchange(long* target, long value)
{
    return InterlockedExchange(target, value);
};
short __declspec (dllexport) __stdcall x64InterlockedExchange16(short* Destination, short ExChange)
{
    return InterlockedExchange16(Destination, ExChange);
};
LONG64 __declspec (dllexport) __stdcall x64InterlockedExchange64(LONG64* target, LONG64 value)
{
    return InterlockedExchange64(target, value);
};
char __declspec (dllexport) __stdcall x64InterlockedExchange8(char* target, char value)
{
    return InterlockedExchange8(target, value);
};

PVOID __declspec (dllexport) __stdcall x64InterlockedExchangePointer(PVOID* target, PVOID value)
{
    return InterlockedExchangePointer(target, value);
};


long __declspec (dllexport) __stdcall x64InterlockedExchangeAdd(long* Addend, long value)
{
    return InterlockedExchangeAdd(Addend, value);
};
LONG64 __declspec (dllexport) __stdcall x64InterlockedExchangeAdd64(LONG64* Addend, LONG64 value)
{
    return InterlockedExchangeAdd64(Addend, value);
};


long __declspec (dllexport) __stdcall x64InterlockedAdd(long* Addend, long value)
{
    return InterlockedAddAcquire(Addend, value);
};
LONG64 __declspec (dllexport) __stdcall x64InterlockedAdd64(LONG64* Addend, LONG64 value)
{
    return InterlockedAdd64(Addend, value);
};

long __declspec (dllexport) __stdcall x64InterlockedAnd(long* Destination, long value)
{
    return InterlockedAnd(Destination, value);
};
char __declspec (dllexport) __stdcall x64InterlockedAnd8(char* Destination, char value)
{
    return InterlockedAnd8(Destination, value);
};
short __declspec (dllexport) __stdcall x64InterlockedAnd16(short* Destination, short value)
{
    return InterlockedAnd16(Destination, value);
};
LONG64 __declspec (dllexport) __stdcall x64InterlockedAnd64(LONG64* Destination, LONG64 value)
{
    return InterlockedAnd64(Destination, value);
};

long __declspec (dllexport) __stdcall x64InterlockedOr(long* Destination, long value)
{
    return InterlockedOr(Destination, value);
};
char __declspec (dllexport) __stdcall x64InterlockedOr8(char* Destination, char value)
{
    return InterlockedOr8(Destination, value);
};
short __declspec (dllexport) __stdcall x64InterlockedOr16(short* Destination, short value)
{
    return InterlockedOr16(Destination, value);
};
LONG64 __declspec (dllexport) __stdcall x64InterlockedOr64(LONG64* Destination, LONG64 value)
{
    return InterlockedOr64(Destination, value);
};

long __declspec (dllexport) __stdcall x64InterlockedXor(long* Destination, long value)
{
    return InterlockedXor(Destination, value);
};
char __declspec (dllexport) __stdcall x64InterlockedXor8(char* Destination, char value)
{
    return InterlockedXor8(Destination, value);
};
short __declspec (dllexport) __stdcall x64InterlockedXor16(short* Destination, short value)
{
    return InterlockedXor16(Destination, value);
};
LONG64 __declspec (dllexport) __stdcall x64InterlockedXor64(LONG64* Destination, LONG64 value)
{
    return InterlockedXor64(Destination, value);
};

LONG __declspec (dllexport) __stdcall x64InterlockedCompareExchange(LONG* Destination, LONG ExChange, LONG Comperand)
{
    return InterlockedCompareExchange(Destination, ExChange, Comperand);
};
short __declspec (dllexport) __stdcall x64InterlockedCompareExchange16(short* Destination, short ExChange, short Comperand)
{
    return InterlockedCompareExchange16(Destination, ExChange, Comperand);
};
LONG64 __declspec (dllexport) __stdcall x64InterlockedCompareExchange64(LONG64* Destination, LONG64 ExChange, LONG64 Comperand)
{
    return InterlockedCompareExchange64(Destination, ExChange, Comperand);
};
BOOLEAN __declspec (dllexport) __stdcall x64InterlockedCompareExchange128(LONG64* Destination, LONG64 ExchangeHigh, LONG64 ExchangeLow, LONG64* ComperandResult)
{
    return InterlockedCompareExchange128(Destination, ExchangeHigh, ExchangeLow, ComperandResult);
};
PVOID __declspec (dllexport) __stdcall x64InterlockedCompareExchangePointer(PVOID* Destination, PVOID Exchange, PVOID Comperand)
{
    return InterlockedCompareExchangePointer(Destination, Exchange, Comperand);
};