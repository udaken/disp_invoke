#pragma once
#include <windows.h>
#include <oleauto.h>
#include <functional>

namespace disp_invoke {

    namespace impl {
        template <class T>
        struct ComValueWrapper
        {
            constexpr ComValueWrapper(T value) : value(value) {}
            T const value;

            constexpr operator T() const {
                return value;
            }
            constexpr operator T& () {
                return value;
            }
        };
    }

    using ComDATE = impl::ComValueWrapper<DATE>;
    using ComSCODE = impl::ComValueWrapper<SCODE>;
    using ComVARIANT_BOOL = impl::ComValueWrapper<VARIANT_BOOL>;

    template <class T>
    struct IDispatchInfo
    {
        virtual LPCTSTR name() const noexcept = 0;
        virtual HRESULT invoke(
            T* self,
            __RPC__in_ecount_full(cArgs) const VARIANTARG* rgvar,
            UINT cArgs,
            _Out_opt_  VARIANT* pVarResult) const noexcept(false) = 0;
        virtual ~IDispatchInfo() {}
    };

    namespace impl {

        class VariantConvertException : public std::exception
        {
        public:
            VariantConvertException(HRESULT hr, signed argNum = -1) : value(hr), argNum(argNum) {}
            HRESULT const value;
            signed const argNum;
        };

        template <unsigned ArgNum>
        struct VariantArgConverter
        {
            const VARIANTARG* const ref;
            constexpr VariantArgConverter(const VARIANTARG& other) : ref{ &other } { }

            constexpr operator CHAR           () const { return V_VT(ref) == VT_I1 ? V_I1(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator SHORT          () const { return V_VT(ref) == VT_I2 ? V_I2(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator LONG           () const { return V_VT(ref) == VT_I4 ? V_I4(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator LONGLONG       () const { return V_VT(ref) == VT_I8 ? V_I8(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator FLOAT          () const { return V_VT(ref) == VT_R4 ? V_R4(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator DOUBLE         () const { return V_VT(ref) == VT_R8 ? V_R8(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator ComVARIANT_BOOL() const { return V_VT(ref) == VT_BOOL ? V_BOOL(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator ComSCODE       () const { return V_VT(ref) == VT_ERROR ? V_ERROR(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator CY             () const { return V_VT(ref) == VT_CY ? V_CY(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator ComDATE        () const { return V_VT(ref) == VT_DATE ? V_DATE(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator BSTR           () const { return V_VT(ref) == VT_BSTR ? V_BSTR(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator IUnknown* () const { return V_VT(ref) == VT_UNKNOWN ? V_UNKNOWN(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator IDispatch* () const { return V_VT(ref) == VT_DISPATCH ? V_DISPATCH(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator SAFEARRAY* () const { return V_VT(ref) == VT_ARRAY ? V_ARRAY(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator CHAR* () const { return V_VT(ref) == (VT_BYREF | VT_I1) ? V_I1REF(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator SHORT* () const { return V_VT(ref) == (VT_BYREF | VT_I2) ? V_I2REF(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator LONG* () const { return V_VT(ref) == (VT_BYREF | VT_I4) ? V_I4REF(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator LONGLONG* () const { return V_VT(ref) == (VT_BYREF | VT_I8) ? V_I8REF(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator FLOAT* () const { return V_VT(ref) == (VT_BYREF | VT_R4) ? V_R4REF(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator DOUBLE* () const { return V_VT(ref) == (VT_BYREF | VT_R8) ? V_R8REF(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator ComVARIANT_BOOL* () const { return V_VT(ref) == (VT_BYREF | VT_BOOL) ? reinterpret_cast<ComVARIANT_BOOL*>(V_BOOLREF(ref)) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator ComSCODE* () const { return V_VT(ref) == (VT_BYREF | VT_ERROR) ? reinterpret_cast<ComSCODE*>(V_ERRORREF(ref)) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator CY* () const { return V_VT(ref) == (VT_BYREF | VT_CY) ? V_CYREF(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator ComDATE* () const { return V_VT(ref) == (VT_BYREF | VT_DATE) ? reinterpret_cast<ComDATE*>(V_DATEREF(ref)) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator BSTR* () const { return V_VT(ref) == (VT_BYREF | VT_BSTR) ? V_BSTRREF(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator IUnknown** () const { return V_VT(ref) == (VT_BYREF | VT_UNKNOWN) ? V_UNKNOWNREF(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator IDispatch** () const { return V_VT(ref) == (VT_BYREF | VT_DISPATCH) ? V_DISPATCHREF(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator SAFEARRAY** () const { return V_VT(ref) == (VT_BYREF | VT_ARRAY) ? V_ARRAYREF(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator VARIANT* () const { return V_VT(ref) == (VT_BYREF | VT_VARIANT) ? V_VARIANTREF(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator PVOID          () const { return V_VT(ref) == VT_BYREF ? V_BYREF(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator BYTE           () const { return V_VT(ref) == VT_UI1 ? V_UI1(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator USHORT         () const { return V_VT(ref) == VT_UI2 ? V_UI2(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator ULONG          () const { return V_VT(ref) == VT_UI4 ? V_UI4(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator ULONGLONG      () const { return V_VT(ref) == VT_UI8 ? V_UI8(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator INT            () const { return V_VT(ref) == VT_INT ? V_INT(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator UINT           () const { return V_VT(ref) == VT_UINT ? V_UINT(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator DECIMAL* () const { return V_VT(ref) == (VT_BYREF | VT_DECIMAL) ? V_DECIMALREF(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator BYTE* () const { return V_VT(ref) == (VT_BYREF | VT_UI1) ? V_UI1REF(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator USHORT* () const { return V_VT(ref) == (VT_BYREF | VT_UI2) ? V_UI2REF(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator ULONG* () const { return V_VT(ref) == (VT_BYREF | VT_UI4) ? V_UI4REF(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator ULONGLONG* () const { return V_VT(ref) == (VT_BYREF | VT_UI8) ? V_UI8REF(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator INT* () const { return V_VT(ref) == (VT_BYREF | VT_INT) ? V_INTREF(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }
            constexpr operator UINT* () const { return V_VT(ref) == (VT_BYREF | VT_UINT) ? V_UINTREF(ref) : throw VariantConvertException(DISP_E_TYPEMISMATCH, ArgNum); }

        };

        struct VariantConverter
        {
            VARIANT* const ref;
            constexpr VariantConverter(VARIANT& other) : ref{ &other } { }

            constexpr operator CHAR& () { return V_VT(ref) = VT_I1, V_I1(ref); }
            constexpr operator SHORT& () { return V_VT(ref) = VT_I2, V_I2(ref); }
            constexpr operator LONG& () { return V_VT(ref) = VT_I4, V_I4(ref); }
            constexpr operator LONGLONG& () { return V_VT(ref) = VT_I8, V_I8(ref); }
            constexpr operator FLOAT& () { return V_VT(ref) = VT_R4, V_R4(ref); }
            constexpr operator DOUBLE& () { return V_VT(ref) = VT_R8, V_R8(ref); }
            operator ComVARIANT_BOOL& () { return V_VT(ref) = VT_BOOL, reinterpret_cast<ComVARIANT_BOOL&>(V_BOOL(ref)); }
            operator ComSCODE& () { return V_VT(ref) = VT_ERROR, reinterpret_cast<ComSCODE&>(V_ERROR(ref)); }
            constexpr operator CY& () { return V_VT(ref) = VT_CY, V_CY(ref); }
            operator ComDATE& () { return V_VT(ref) = VT_DATE, reinterpret_cast<ComDATE&>(V_DATE(ref)); }
            constexpr operator BSTR& () { return V_VT(ref) = VT_BSTR, V_BSTR(ref); }
            constexpr operator IUnknown*& () { return V_VT(ref) = VT_UNKNOWN, V_UNKNOWN(ref); }
            constexpr operator IDispatch*& () { return V_VT(ref) = VT_DISPATCH, V_DISPATCH(ref); }
            constexpr operator SAFEARRAY*& () { return V_VT(ref) = VT_ARRAY, V_ARRAY(ref); }
            constexpr operator CHAR*& () { return V_VT(ref) = (VT_BYREF | VT_I1), V_I1REF(ref); }
            constexpr operator SHORT*& () { return V_VT(ref) = (VT_BYREF | VT_I2), V_I2REF(ref); }
            constexpr operator LONG*& () { return V_VT(ref) = (VT_BYREF | VT_I4), V_I4REF(ref); }
            constexpr operator LONGLONG*& () { return V_VT(ref) = (VT_BYREF | VT_I8), V_I8REF(ref); }
            constexpr operator FLOAT*& () { return V_VT(ref) = (VT_BYREF | VT_R4), V_R4REF(ref); }
            constexpr operator DOUBLE*& () { return V_VT(ref) = (VT_BYREF | VT_R8), V_R8REF(ref); }
            operator ComVARIANT_BOOL*& () { return V_VT(ref) = (VT_BYREF | VT_BOOL), reinterpret_cast<ComVARIANT_BOOL*&>(V_BOOLREF(ref)); }
            operator ComSCODE*& () { return V_VT(ref) = (VT_BYREF | VT_ERROR), reinterpret_cast<ComSCODE*&>(V_ERRORREF(ref)); }
            constexpr operator CY*& () { return V_VT(ref) = (VT_BYREF | VT_CY), V_CYREF(ref); }
            operator ComDATE*& () { return V_VT(ref) = (VT_BYREF | VT_DATE), reinterpret_cast<ComDATE*&>(V_DATEREF(ref)); }
            constexpr operator BSTR*& () { return V_VT(ref) = (VT_BYREF | VT_BSTR), V_BSTRREF(ref); }
            constexpr operator IUnknown**& () { return V_VT(ref) = (VT_BYREF | VT_UNKNOWN), V_UNKNOWNREF(ref); }
            constexpr operator IDispatch**& () { return V_VT(ref) = (VT_BYREF | VT_DISPATCH), V_DISPATCHREF(ref); }
            constexpr operator SAFEARRAY**& () { return V_VT(ref) = (VT_BYREF | VT_ARRAY), V_ARRAYREF(ref); }
            constexpr operator VARIANT*& () { return V_VT(ref) = (VT_BYREF | VT_VARIANT), V_VARIANTREF(ref); }
            constexpr operator PVOID& () { return V_VT(ref) = VT_BYREF, V_BYREF(ref); }
            constexpr operator BYTE& () { return V_VT(ref) = VT_UI1, V_UI1(ref); }
            constexpr operator USHORT& () { return V_VT(ref) = VT_UI2, V_UI2(ref); }
            constexpr operator ULONG& () { return V_VT(ref) = VT_UI4, V_UI4(ref); }
            constexpr operator ULONGLONG& () { return V_VT(ref) = VT_UI8, V_UI8(ref); }
            constexpr operator INT& () { return V_VT(ref) = VT_INT, V_INT(ref); }
            constexpr operator UINT& () { return V_VT(ref) = VT_UINT, V_UINT(ref); }
            constexpr operator DECIMAL*& () { return V_VT(ref) = (VT_BYREF | VT_DECIMAL), V_DECIMALREF(ref); }
            constexpr operator BYTE*& () { return V_VT(ref) = (VT_BYREF | VT_UI1), V_UI1REF(ref); }
            constexpr operator USHORT*& () { return V_VT(ref) = (VT_BYREF | VT_UI2), V_UI2REF(ref); }
            constexpr operator ULONG*& () { return V_VT(ref) = (VT_BYREF | VT_UI4), V_UI4REF(ref); }
            constexpr operator ULONGLONG*& () { return V_VT(ref) = (VT_BYREF | VT_UI8), V_UI8REF(ref); }
            constexpr operator INT*& () { return V_VT(ref) = (VT_BYREF | VT_INT), V_INTREF(ref); }
            constexpr operator UINT*& () { return V_VT(ref) = (VT_BYREF | VT_UINT), V_UINTREF(ref); }
        };

        template <class T>
        constexpr bool false_v = false;

        template <class T, class Callable, class ... Args>
        class MethodInfoImpl final : public IDispatchInfo<T>
        {
            LPCTSTR m_func_name;
            Callable const callable;
        public:
            constexpr MethodInfoImpl(LPCTSTR func_name, Callable&& c)
                : m_func_name(func_name), callable(c) {
            }

            constexpr LPCTSTR name() const noexcept
            {
                return m_func_name;
            }

            constexpr HRESULT invoke(T* self, const VARIANTARG* args, UINT cArgs, VARIANT* pVarResult) const override {
                if (sizeof...(Args) > cArgs)
                    return DISP_E_BADPARAMCOUNT;

                if (pVarResult)
                    V_VT(pVarResult) = VT_EMPTY;

                if constexpr (sizeof...(Args) == 0)
                {
                    UNREFERENCED_PARAMETER(args);
                    return callable(self);
                }
                else if constexpr (sizeof...(Args) == 1)
                {
                    return callable(self,
                        VariantArgConverter<0>(args[0]));
                }
                else if constexpr (sizeof...(Args) == 2)
                {
                    return callable(self,
                        VariantArgConverter<0>(args[0]),
                        VariantArgConverter<1>(args[1]));
                }
                else if constexpr (sizeof...(Args) == 3)
                {
                    return callable(self,
                        VariantArgConverter<0>(args[0]),
                        VariantArgConverter<1>(args[1]),
                        VariantArgConverter<2>(args[2]));
                }
                else if constexpr (sizeof...(Args) == 4)
                {
                    return callable(self,
                        VariantArgConverter<0>(args[0]),
                        VariantArgConverter<1>(args[1]),
                        VariantArgConverter<2>(args[2]),
                        VariantArgConverter<3>(args[3]));
                }
                else
                {
                    static_assert(false_v<Args...>);
                }
            }
        };
        template <class T, class Callable, class R, class ... Args>
        class MethodInfoImpl_with_retval final : public IDispatchInfo<T>
        {
            LPCTSTR m_func_name;
            Callable const callable;
        public:
            constexpr MethodInfoImpl_with_retval(LPCTSTR func_name, Callable&& c)
                : m_func_name(func_name), callable(c) {
            }

            constexpr LPCTSTR name() const noexcept
            {
                return m_func_name;
            }

            constexpr HRESULT invoke(T* self, const VARIANTARG* args, UINT cArgs, VARIANT* pVarResult) const override {
                if (sizeof...(Args) > cArgs)
                    return DISP_E_BADPARAMCOUNT;

                R* retval = nullptr;
                if (pVarResult)
                {
                    retval = &(static_cast<R&>(VariantConverter(*pVarResult)));
                }

                if constexpr (sizeof...(Args) == 0)
                {
                    UNREFERENCED_PARAMETER(args);
                    return callable(self, retval);
                }
                else if constexpr (sizeof...(Args) == 1)
                {
                    return callable(self,
                        VariantArgConverter<0>(args[0]),
                        retval);
                }
                else if constexpr (sizeof...(Args) == 2)
                {
                    return callable(self,
                        VariantArgConverter<0>(args[0]),
                        VariantArgConverter<1>(args[1]),
                        retval);
                }
                else if constexpr (sizeof...(Args) == 3)
                {
                    return callable(self,
                        VariantArgConverter<0>(args[0]),
                        VariantArgConverter<1>(args[1]),
                        VariantArgConverter<1>(args[2]),
                        retval);
                }
                else if constexpr (sizeof...(Args) == 4)
                {
                    return callable(self,
                        VariantArgConverter<0>(args[0]),
                        VariantArgConverter<1>(args[1]),
                        VariantArgConverter<1>(args[2]),
                        VariantArgConverter<1>(args[3]),
                        retval);
                }
                else
                {
                    static_assert(false_v<Args...>);
                }
            }
        };

    }

    template <class T, class ... Args>
    inline constexpr std::shared_ptr<const IDispatchInfo<T>> make_method_info(
        LPCTSTR func_name, HRESULT(T::* pm)(Args...)) {
        return std::shared_ptr< IDispatchInfo<T>>(new impl::MethodInfoImpl<T, decltype(std::mem_fn(pm)), Args...>(func_name, std::mem_fn(pm)));
    }

    template <class T, class R>
    inline constexpr std::shared_ptr<const IDispatchInfo<T>> make_method_info_has_retval(
        LPCTSTR func_name, HRESULT(T::* pm)(R*)) {
        return std::shared_ptr< IDispatchInfo<T>>(new impl::MethodInfoImpl_with_retval<T, decltype(std::mem_fn(pm)), R>(func_name, std::mem_fn(pm)));
    }

    template <class T, class R, class Args1>
    inline constexpr std::shared_ptr<const IDispatchInfo<T>> make_method_info_has_retval(
        LPCTSTR func_name, HRESULT(T::* pm)(Args1, R*)) {
        return std::shared_ptr< IDispatchInfo<T>>(new impl::MethodInfoImpl_with_retval<T, decltype(std::mem_fn(pm)), R, Args1>(func_name, std::mem_fn(pm)));
    }

    template <class T, class R, class Args1, class Args2>
    inline constexpr std::shared_ptr<const IDispatchInfo<T>> make_method_info_has_retval(
        LPCTSTR func_name, HRESULT(T::* pm)(Args1, Args2, R*)) {
        return std::shared_ptr< IDispatchInfo<T>>(new impl::MethodInfoImpl_with_retval<T, decltype(std::mem_fn(pm)), R, Args1, Args2>(func_name, std::mem_fn(pm)));
    }

    template <class T, class R, class Args1, class Args2, class Args3>
    inline constexpr std::shared_ptr<const IDispatchInfo<T>> make_method_info_has_retval(
        LPCTSTR func_name, HRESULT(T::* pm)(Args1, Args2, Args3, R*)) {
        return std::shared_ptr< IDispatchInfo<T>>(new impl::MethodInfoImpl_with_retval<T, decltype(std::mem_fn(pm)), R, Args1, Args2, Args3>(func_name, std::mem_fn(pm)));
    }

    template <class T, class R, class Args1, class Args2, class Args3, class Args4>
    inline constexpr std::shared_ptr<const IDispatchInfo<T>> make_method_info_has_retval(
        LPCTSTR func_name, HRESULT(T::* pm)(Args1, Args2, Args3, Args4, R*)) {
        return std::shared_ptr< IDispatchInfo<T>>(new impl::MethodInfoImpl_with_retval<T, decltype(std::mem_fn(pm)), R, Args1, Args2, Args3, Args4>(func_name, std::mem_fn(pm)));
    }

    template <class T>
    inline std::unordered_map<DISPID, std::shared_ptr<const IDispatchInfo<T>>> make_dispatch_info_table(std::initializer_list<std::shared_ptr<const IDispatchInfo<T>>>&& args) {
        std::unordered_map<DISPID, std::shared_ptr<const IDispatchInfo<T>>> map{};
        DISPID id{};
        for (auto&& i : args)
        {
            map.emplace(id++, i);
        }
        return map;
    }
    template <class T>
    struct DECLSPEC_NOVTABLE DispInvokeSupport abstract : IDispatch
    {
        IFACEMETHODIMP GetTypeInfoCount(
            /* [out] */ __RPC__out UINT * pctinfo) noexcept override
        {
            *pctinfo = 0;
            return S_OK;
        }

        IFACEMETHODIMP GetTypeInfo(
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ __RPC__deref_out_opt ITypeInfo * *ppTInfo) noexcept override
        {
            UNREFERENCED_PARAMETER(iTInfo);
            UNREFERENCED_PARAMETER(lcid);
            *ppTInfo = nullptr;
            return E_NOTIMPL;
        }

        IFACEMETHODIMP GetIDsOfNames(
            /* [in] */ __RPC__in REFIID riid,
            /* [size_is][in] */ __RPC__in_ecount_full(cNames) LPOLESTR * rgszNames,
            /* [range][in] */ __RPC__in_range(0, 16384) UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ __RPC__out_ecount_full(cNames) DISPID * rgDispId) noexcept override
        {
            using namespace impl;

            UNREFERENCED_PARAMETER(lcid);
            if ((riid != IID_NULL))
                return E_INVALIDARG;

            HRESULT hr = S_OK;
            for (UINT i = 0; i < cNames; i++)
            {
                rgDispId[i] = DISPID_UNKNOWN;
                for (auto& itr : T::get_method_info_table())
                {
                    if (wcscmp(rgszNames[i], itr.second->name()) == 0)
                    {
                        rgDispId[i] = itr.first;
                    }
                }
                if (rgDispId[i] == DISPID_UNKNOWN)
                    hr = DISP_E_UNKNOWNNAME;
            }
            return S_OK;
        }

        IFACEMETHODIMP Invoke(
            /* [annotation][in] */
            _In_  DISPID dispIdMember,
            /* [annotation][in] */
            _In_  REFIID riid,
            /* [annotation][in] */
            _In_  LCID lcid,
            /* [annotation][in] */
            _In_  WORD wFlags,
            /* [annotation][out][in] */
            _In_  DISPPARAMS * pDispParams,
            /* [annotation][out] */
            _Out_opt_  VARIANT * pVarResult,
            /* [annotation][out] */
            _Out_opt_  EXCEPINFO * pExcepInfo,
            /* [annotation][out] */
            _Out_opt_  UINT * puArgErr) noexcept override
        {
            using namespace impl;

            UNREFERENCED_PARAMETER(lcid);
            if ((wFlags & DISPATCH_METHOD) == 0)
                return E_INVALIDARG;

            if ((riid != IID_NULL))
                return E_INVALIDARG;

            // named argument is not supported
            if (pDispParams->cNamedArgs != 0)
                return DISP_E_NONAMEDARGS;


            auto itr{ T::get_method_info_table().find(dispIdMember)};
            if (itr == T::get_method_info_table().end())
            {
                return DISP_E_MEMBERNOTFOUND;
            }

            if (pVarResult) {
                *pVarResult = {};
            }

            if (pExcepInfo) {
                *pExcepInfo = {};
            }

            if (puArgErr) {
                *puArgErr = {};
            }

            try
            {
                HRESULT hr = itr->second->invoke(static_cast<T*>(this), pDispParams->rgvarg, pDispParams->cArgs, pVarResult);

                if (FAILED(hr))
                {
                    if (pExcepInfo) {
                        pExcepInfo->scode = hr;
                    }
                    return DISP_E_EXCEPTION;
                }
                return hr;
            }
            catch (const VariantConvertException& ex)
            {
            if (ex.value == DISP_E_TYPEMISMATCH && puArgErr)
            *puArgErr = ex.argNum;
            return ex.value;
            }
        }
    };
}
