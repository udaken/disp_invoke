#include <iostream>
#include <assert.h>
#define NOMINMAX
#define STRICT
#define WIN32_LEAN_AND_MEAN
#include <windows.h>

#include "disp_invoke.hpp"

using namespace disp_invoke;

class SomeClass final : public DispInvokeSupport<SomeClass>
{
public:
    STDMETHODIMP f()
    {
        return 7;
    };
    STDMETHODIMP f1(int param, BSTR name, BSTR* retval)
    {
        WCHAR buf[512];
        wsprintf(buf, L"Hello %s", name);
        *retval = SysAllocString(buf);
        return param;
    };

    IFACEMETHODIMP QueryInterface(
        /* [in] */ REFIID riid,
        /* [iid_is][out] */ _COM_Outptr_ void __RPC_FAR* __RPC_FAR* ppvObject) override
    {
        return S_OK;
    }

    IFACEMETHODIMP_(ULONG) AddRef(void) override
    {
        return 0;
    }

    IFACEMETHODIMP_(ULONG) Release(void) override
    {
        return 0;
    }

public:
    static auto& get_method_info_table() { return s_table; }
    inline static const std::unordered_map<DISPID, std::shared_ptr<const IDispatchInfo<SomeClass>>> s_table = 
        make_dispatch_info_table({
            make_method_info(TEXT("f"), &SomeClass::f),
            make_method_info_has_retval(TEXT("f1"), &SomeClass::f1),
        });

};

int main()
{
    HRESULT hr;
    SomeClass s;
    DISPID dispId;
    OLECHAR name[] = L"f1";
    LPOLESTR names[] = { name };
    hr = s.GetIDsOfNames(IID_NULL, names, ARRAYSIZE(names), 0, &dispId);
    assert(SUCCEEDED(hr));

    VARIANTARG args[2]{};
    V_VT(&args[0]) = VT_INT; V_INT(&args[0]) = 99;
    V_VT(&args[1]) = VT_BSTR; V_BSTR(&args[1]) = SysAllocString(L"bob");

    DISPPARAMS dp{};
    dp.cArgs = ARRAYSIZE(args);
    dp.rgvarg = args;

    VARIANT varResult;
    UINT argErr;
    EXCEPINFO exceptionInfo;
    hr = s.Invoke(dispId, IID_NULL, 0, DISPATCH_METHOD, &dp, &varResult, &exceptionInfo, &argErr);
    assert(SUCCEEDED(hr));

    std::wcout << (BSTR)impl::VariantArgConverter<-1>(varResult);
}
