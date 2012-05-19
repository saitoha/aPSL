
#include <ActivScp.h>
#include <ComCat.h>
#include <comdef.h>
#include <vector>

#include "PSL/PSL.h"

HINSTANCE hInst;

#define APSL_ASSERT(x)
#define APSL_TRACE(x)
#define PACKAGE_NAME "aPSL"
#define IID_APSL "{B7BCEFC5-FD47-4986-B418-A6686F9760CC}"


namespace aPSL { namespace util {
    
    //////////////////////////////////////////////////////////////////////////
    //
    //  @class scoped_lock
    //  @brief
    //
    class scoped_lock
    {
    public:
        explicit scoped_lock(CRITICAL_SECTION &cs) throw()
        : pcs_(&cs)
        {
            EnterCriticalSection(pcs_);
        }

        ~scoped_lock() throw()
        {
            LeaveCriticalSection(pcs_);
        }

        CRITICAL_SECTION *pcs_;
    };

    //////////////////////////////////////////////////////////////////////////
    //
    //  @class critical_section
    //
    struct critical_section
    : public CRITICAL_SECTION
    {
        critical_section() throw()
        {
            InitializeCriticalSection(this);
        }

        ~critical_section() throw()
        {
            DeleteCriticalSection(this);
        }
    };

} } // namespace aPSL::util

namespace aPSL {

    VARIANT variable_to_variant(PSL::variable const& v);
    PSL::variable * variant_to_variable(VARIANT const& v);

    class activex_object;


    //////////////////////////////////////////////////////////////////////////
    //
    //  @class com_callable_wrapper
    //
    struct com_callable_wrapper
    : IDispatch
    {
        explicit com_callable_wrapper(PSL::variable * primitive) throw()
        : m_count(0)
        , primitive_(primitive)
        {
        }

        ~com_callable_wrapper() throw()
        {
        }

    // IUnknown implementation
        HRESULT STDMETHODCALLTYPE QueryInterface(
                /* [in] */ REFIID riid,
                /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject) throw()
        {
	        if (!ppvObject)
	            return E_POINTER;
	        if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch))
	            return *ppvObject = this, S_OK;
	        return E_NOINTERFACE;
        }

        ULONG STDMETHODCALLTYPE AddRef() throw()
        {
            util::scoped_lock lock(critical_section_);
            return ++ m_count;
        }

        ULONG STDMETHODCALLTYPE Release() throw()
        {
            util::scoped_lock lock(critical_section_);
            if (m_count < 1)
                delete this;
            return -- m_count;
        }

    // IDispatch implementation
        STDMETHOD(GetTypeInfoCount)(UINT*) throw()
        {
            return E_NOTIMPL;
        }

        STDMETHOD(GetTypeInfo)(UINT, LCID, ITypeInfo**) throw()
        {
            return E_NOTIMPL;
        }

        STDMETHOD(GetIDsOfNames)(
            REFIID,
            LPOLESTR* rgszNames,
            UINT cNames,
            LCID,
            DISPID* rgdispid) throw()
        {
            util::scoped_lock lock(critical_section_);
            APSL_ASSERT(cNames == 1);
            key_.push_back(*rgszNames);
            *rgdispid = DISPID(key_.size());
            return S_OK;
        }

        STDMETHOD(Invoke)(
            DISPID dispidMember,
            REFIID,
            LCID,
            WORD wFlags,
            DISPPARAMS* pdispparams,
            VARIANT* pvarResult,
            EXCEPINFO* pexcepinfo,
            UINT* puArgErr)
        {
            util::scoped_lock lock(critical_section_);
            switch (wFlags)
            {
            case DISPATCH_METHOD:
                return invoke_method(pdispparams, pvarResult);
            case DISPATCH_PROPERTYGET:
                return invoke_propertyget(dispidMember, pvarResult);
            default:
                APSL_ASSERT(!"com_callable_wrapper::Invoke");
			    return E_UNEXPECTED;
            }
            __assume(0);
        }
    private:
        HRESULT invoke_method(DISPPARAMS* pdispparams, VARIANT* pvarResult) const throw()
        {
            try {
                PSL::variable arg(PSL::variable::RARRAY);
                for (UINT i = 0; i < pdispparams->cArgs; ++i)
                   arg.push(variant_to_variable(pdispparams->rgvarg[pdispparams->cArgs - i - 1]));
                *pvarResult = variable_to_variant(primitive_->operator ()(arg));
            }
            catch (...) {
                APSL_ASSERT(0);
                *pvarResult = _variant_t();
                return E_UNEXPECTED;
            }
            return S_OK;
        }

        HRESULT invoke_propertyget(DISPID dispidMember, VARIANT* pvarResult) const throw()
        {
            try {
                if (dispidMember == 0)
                {
                    const char *str = primitive_->operator PSL::string();
                    BSTR olestr = _bstr_t(str);
                    *pvarResult = _variant_t(olestr);
                }
                else
                {
                    *pvarResult = variable_to_variant(primitive_[(key_[dispidMember - 1])]);
                }
                return S_OK;
            }
            catch (...) {
                *pvarResult = _variant_t();
                return E_UNEXPECTED;
            }
        }

    private:
        ULONG m_count;
        PSL::variable * primitive_;
        util::critical_section critical_section_;
        std::vector<PSL::string> key_;
    };

    VARIANT variable_to_variant(PSL::variable const& v)
    {
        switch (v.type()) {

        case PSL::variable::NIL:
            {
                VARIANT v;
                v.vt = VT_NULL;
                return v;
            }
            break;

		case PSL::variable::INT:
            return _variant_t(v.operator int());

		case PSL::variable::HEX:
            return _variant_t(v.operator unsigned char());

		case PSL::variable::FLOAT:
            return _variant_t(v.operator double());

		case PSL::variable::STRING:
            return _variant_t(v.operator char const *());

		case PSL::variable::POINTER:
            return _variant_t((LPDISPATCH)(void *)v);

		case PSL::variable::RARRAY:
            APSL_ASSERT(0);
            // TODO:
            //x = new vRArray();
            break;

		case PSL::variable::THREAD:
            APSL_ASSERT(0);
            // TODO:
            //x = new vThread();
            //return reinterpret_cast<runtime_callable_wrapper const&>(value).get_dispatch();
            break;

		default:
            {
                // TODO: exception handling
                IDispatch *pdisp = new aPSL::com_callable_wrapper(v);
                pdisp->AddRef();
                return _variant_t(pdisp);
            }
        }
        __assume(0);
    }

    //////////////////////////////////////////////////////////////////////////
    //
    //  @class runtime_callable_wrapper
    //
    class runtime_callable_wrapper : public PSL::variable
    {
    public:
        explicit runtime_callable_wrapper(IDispatch *pDispatch, DISPID dispid) throw()
        : m_pDispatch(pDispatch)
        , m_dispid(dispid)
        {
            m_pDispatch->AddRef();
        }

        ~runtime_callable_wrapper() throw()
        {
            m_pDispatch->Release();
        }

        PSL::variable * __stdcall call__(PSL::variable& /*this_arg*/, PSL::variable& arguments)
        {
            return call_impl(arguments);
        }

        PSL::variable * __stdcall get_value__()
        {
            return get_value_impl();
        }

        PSL::variable * __stdcall assign__(PSL::variable& rhs)
        {
            return assign_impl(rhs);
        }

        IDispatch *get_dispatch() const
        {
            return m_pDispatch;
        }

    private:
        PSL::variable * __stdcall call_impl(PSL::variable& arguments)
        {
            VARIANT result = {VT_EMPTY};
            EXCEPINFO excepinfo = {0};
            UINT argerr = 0;
            std::vector<VARIANTARG> variant_arg;
            size_t length = arguments.length();
            for (size_t i = 0; i < length; ++i)
                variant_arg.push_back(variable_to_variant(arguments[length - i - 1]));
            DISPPARAMS params
                = {length > 0 ? &*variant_arg.begin(): NULL, NULL, length, 0};
            HRESULT hr = m_pDispatch->Invoke(
                m_dispid, IID_NULL, LOCALE_USER_DEFAULT,
                DISPATCH_METHOD, &params, &result, &excepinfo, &argerr);
            if (DISP_E_EXCEPTION == hr)
                APSL_ASSERT(0);
            if (FAILED(hr))
                APSL_ASSERT(0);
            return variant_to_variable(result);
        }

        PSL::variable * get_value_impl()
        {
            VARIANT result = {VT_EMPTY};
            EXCEPINFO excepinfo = {0};
            UINT argerr = 0;
            DISPPARAMS params = {NULL, NULL, 0, 0};
            HRESULT hr = m_pDispatch->Invoke(
                m_dispid, IID_NULL, LOCALE_USER_DEFAULT,
                DISPATCH_PROPERTYGET, &params, &result, &excepinfo, &argerr);
            if (FAILED(hr))
                APSL_ASSERT(0);
            return variant_to_variable(result);
        }

        PSL::variable * assign_impl(PSL::variable& rhs)
        {
            VARIANT result = {VT_EMPTY};
            EXCEPINFO excepinfo = {0};
            UINT argerr = 0;
            VARIANT value = variable_to_variant(rhs);
            DISPID dispid = DISPID_PROPERTYPUT;
            DISPPARAMS params = {&value, &dispid, 1, 1};
            HRESULT hr = m_pDispatch->Invoke(
                m_dispid, IID_NULL, LOCALE_USER_DEFAULT,
                DISPATCH_PROPERTYPUT, &params, &result, &excepinfo, &argerr);
            if (FAILED(hr))
                APSL_ASSERT(0);
            return rhs;
        }

    private:
        IDispatch *m_pDispatch;
        DISPID m_dispid;
    };


    //////////////////////////////////////////////////////////////////////////
    //
    //  @class activex_object
    //
    class activex_object: public PSL::variable
    {
    public:
        explicit activex_object(LPDISPATCH pdisp) throw()
        : PSL::variable(pdisp)
        , m_pDispatch(pdisp)
        {
            APSL_ASSERT (NULL != m_pDispatch);
            m_pDispatch->AddRef();
        }

        virtual ~activex_object() throw()
        {
            m_pDispatch->Release();
        }

        PSL::variable * __stdcall get__(PSL::string const& key)
        {
            return get_impl(key);
        }

        void __stdcall put__(PSL::string const& key, PSL::variable *rhs)
        {
            put_impl(key, rhs);
        }

    private:
        PSL::variable *get_impl(PSL::string const& key)
        {
            LPOLESTR rgszNames = _bstr_t(key.c_str());
            DISPID rgDispid = 0;
            HRESULT hr = S_OK;
            hr = m_pDispatch->GetIDsOfNames(
                IID_NULL, &rgszNames, 1, LOCALE_USER_DEFAULT, &rgDispid);
            if (hr == DISP_E_UNKNOWNNAME)
                return this->operator [] (key);
            if (SUCCEEDED(hr))
                return new runtime_callable_wrapper(m_pDispatch, rgDispid);
//            _com_error const e(hr);
//            fwprintf(stderr, L"com error: code=%d message=%s\n", e.Error(), e.ErrorMessage());
            throw std::runtime_error("com error: ");
        }

     private:
        void put_impl(PSL::string const& key, PSL::variable *rhs)
        {
            LPOLESTR rgszNames = _bstr_t(key.c_str());
            DISPID rgDispid = 0;
            HRESULT hr = m_pDispatch->GetIDsOfNames(
                IID_NULL, &rgszNames, 1, LOCALE_USER_DEFAULT, &rgDispid);
            if (hr != S_OK)
                throw (hr);
            VARIANT result = {VT_EMPTY};
            EXCEPINFO excepinfo = {0};
            UINT argerr = 0;
            VARIANT value = variable_to_variant(rhs);
            DISPPARAMS params = {&value, &rgDispid, 1, 1};
            hr = m_pDispatch->Invoke(
                rgDispid, IID_NULL, LOCALE_USER_DEFAULT,
                DISPATCH_PROPERTYPUT, &params, &result, &excepinfo, &argerr);
            if (FAILED(hr))
                throw (hr);
            hr = m_pDispatch->GetIDsOfNames(
                IID_NULL, &rgszNames, 1, LOCALE_USER_DEFAULT, &rgDispid);
            if (hr != S_OK)
                throw (hr);
        }

     private:
        IDispatch *m_pDispatch;
    };


    PSL::variable * variant_to_variable(VARIANT const& v)
    {
        switch (v.vt & VT_TYPEMASK)
        {
        case VT_DISPATCH:
            return new aPSL::activex_object(v.vt & VT_BYREF ? *v.ppdispVal: v.pdispVal);
        case VT_EMPTY:
            return new PSL::variable; // NIL
        case VT_NULL:
            return new PSL::variable; // NIL
        case VT_BOOL:
            if ((v.vt & VT_BYREF ? *v.pboolVal: v.boolVal) == VARIANT_TRUE)
                return new PSL::variable(1);
            else
                return new PSL::variable(0);
        case VT_BSTR:
            {
                BSTR bstr = v.vt & VT_BYREF ? *v.pbstrVal: v.bstrVal;
                if (!bstr)
                    return new PSL::variable("");
                UINT length = ::SysStringByteLen(bstr);
                const char *str = _bstr_t(bstr, bstr + length);
                PSL::variable *result = new PSL::variable(PSL::string(str));
                ::SysFreeString(bstr);
                return result;
            }
        case VT_I4:
            return new PSL::variable(v.vt & VT_BYREF ? *v.plVal: v.lVal);
        case VT_UI4:
            return new PSL::variable(v.vt & VT_BYREF ? *v.pulVal: v.ulVal);
        case VT_R8:
            return new PSL::variable(v.vt & VT_BYREF ? *v.pdblVal: v.dblVal);
        case VT_ARRAY:
        case VT_VECTOR:
            APSL_ASSERT(0);
        default:
            APSL_ASSERT(0);
        }
        __assume(0);
    }

} // namespace aPSL


namespace aPSL {

    //////////////////////////////////////////////////////////////////////
    //
    //  @class scriptsite_object
    //
    class scriptsite_object 
    : public PSL::variable
    {
    public:
        explicit scriptsite_object(IActiveScriptSite *pActiveScriptSite)
        : m_pActiveScriptSite(pActiveScriptSite)
        {
            APSL_ASSERT(NULL != m_pActiveScriptSite);
            m_pActiveScriptSite->AddRef();
        }
    
        virtual ~scriptsite_object() throw()
        {
            APSL_ASSERT(NULL != m_pActiveScriptSite);
            m_pActiveScriptSite->Release();
        }
    
        PSL::variable *get__(LPCOLESTR key)
        {
            LPDISPATCH pdisp = NULL;
            HRESULT hr = get_member(key, &pdisp);
            PSL::variable *result;
            if (FAILED(hr) || NULL == pdisp)
                result = new PSL::variable;
            else
                result = new aPSL::activex_object(pdisp);
            result->ref();
            return result;
        }
    
        HRESULT get_member(LPCOLESTR key, LPDISPATCH *ppdisp)
        {
            LPOLESTR rgszNames = const_cast<LPOLESTR>(key);
            DISPID rgDispid = 0;
            HRESULT hr = S_OK;
            IUnknown *pUnkown = NULL;
            APSL_ASSERT(0 != m_pActiveScriptSite);
            hr = m_pActiveScriptSite->GetItemInfo(
                rgszNames, SCRIPTINFO_IUNKNOWN, &pUnkown, NULL);
            if (FAILED(S_OK))
                return hr;
            return pUnkown->QueryInterface(
                IID_IDispatch, reinterpret_cast<LPVOID*>(ppdisp));
        }
    
    private:
        IActiveScriptSite *m_pActiveScriptSite;
    };

    //////////////////////////////////////////////////////////////////////
    //
    //  @class script_engine
    //
    class script_engine
    {
    public:
        script_engine()
        {
        }

        void eval(const char *text)
        {
            vm.LoadString(text);
            vm.Run();
        }

        void put__(const PSL::string& pstrName, const PSL::variable& v)
        {
            vm.add(pstrName, v);
        }

    private:
        PSL::PSLVM vm;
    };


    //////////////////////////////////////////////////////////////////////////////
    //
    //  @class     com_module
    //
    class com_module
    {
    public:
        com_module() throw()
        : count_(0)
        {
        }
    
        int increment() throw()
        {
            if (INT_MAX <= count_)
                return -1;
            InterlockedIncrement(&count_);
            return 0;
        }
    
        int decrement() throw()
        {
            APSL_ASSERT(count_ > 0);
            if (count_ <= 0)
                return -1;
            InterlockedDecrement(&count_);
            return 0;
        }
    
        int can_unload_now() const throw()
        {
            APSL_ASSERT(count_ >= 0);
            return count_ == 0 ? 0 : -1;
        }
    
    private:
        long count_;
    
    };

} // namespace aPSL

/* ---- object implementation ---- */

///////////////////////////////////////////////////////////////////////////
//
// @struct INTERFACE_ENTRY
//
struct INTERFACE_ENTRY
{
    IID const * p_iid;
    IUnknown *interface_pointer;
};

///////////////////////////////////////////////////////////////////////////
//
// @class IUnknownImpl
//
template <class T>
class IUnknownImpl 
: T
{
public:
    IUnknownImpl() throw()
    : m_Count(0)
    {
    }

    STDMETHODIMP_(ULONG) AddRef(VOID)
    {
        return ++m_Count;
    }

    STDMETHODIMP_(ULONG) Release(VOID)
    {
        return --m_Count;
    };

    STDMETHOD(QueryInterface)(REFIID iid, LPVOID *ppv)
    {
        INTERFACE_ENTRY const* p_entry;
        for (p_entry = T::GetInterfaceMap(); NULL != p_entry->p_iid; ++p_entry)
            if (IsEqualGUID(iid, *p_entry->p_iid))
                return *ppv = p_entry->interface_pointer, S_OK;
        return E_NOINTERFACE;
    }

private:
    ULONG m_Count;

};

///////////////////////////////////////////////////////////////////////////
//
// @class IActiveScriptImpl
//
template <class T>
class __declspec(novtable) IActiveScriptImpl
: public IActiveScript
{
    public:
        IActiveScriptImpl() throw()
        : m_p_script_engine(NULL)
        , m_p_scriptsite_object(NULL)
        {
        }

        ~IActiveScriptImpl() throw()
        {
            if (m_p_scriptsite_object)
                delete m_p_scriptsite_object;
        }

        STDMETHOD(SetScriptSite)(IActiveScriptSite *pass)
        {
            APSL_TRACE ("IActiveScript::SetScriptSite");
            m_ActiveScriptSite = pass;
            m_p_scriptsite_object = new aPSL::scriptsite_object(m_ActiveScriptSite);
            if (NULL == m_p_scriptsite_object)
                return E_OUTOFMEMORY;
            m_script_state = SCRIPTSTATE_INITIALIZED;
            PSL::variable *window = m_p_scriptsite_object->get__(L"window");
            
            m_p_script_engine = new aPSL::script_engine();
            m_p_script_engine->put__("window", *window);
            return S_OK;
        }

        STDMETHOD(GetScriptSite)(REFIID riid, LPVOID *ppvObject)
        {
            APSL_TRACE ("IActiveScript::GetScriptSite");
            return m_ActiveScriptSite->QueryInterface(riid, ppvObject);
        }

        STDMETHOD(SetScriptState)(SCRIPTSTATE ss)
        {
            APSL_TRACE ("IActiveScript::SetScriptState");
            if (m_script_state = ss)
                return S_FALSE;
	        if (SCRIPTSTATE_UNINITIALIZED != ss)
		        m_ActiveScriptSite->OnStateChange(ss);
            m_script_state = ss;
            return S_OK;
        }

        STDMETHOD(GetScriptState)(SCRIPTSTATE *pssState)
        {
            APSL_TRACE ("IActiveScript::GetScriptState");
            return *pssState = m_script_state, S_OK;
        }

        STDMETHOD(Close)(VOID)
        {
            APSL_TRACE ("IActiveScript::Close");
            return S_OK;
        };

        STDMETHOD(AddNamedItem)(LPCOLESTR pstrName, DWORD dwFlags)
        {
            APSL_TRACE ("IActiveScript::AddNamedItem");
            if (!m_p_script_engine)
                return E_POINTER;
            if (!m_p_scriptsite_object)
                return E_POINTER;
            
            PSL::variable *scriptsite = m_p_scriptsite_object->get__(pstrName);
            m_p_script_engine->put__(pstrName, scriptsite);
            return S_OK;
        }

        STDMETHOD(AddTypeLib)(REFGUID, DWORD, DWORD, DWORD)
        {
            APSL_TRACE ("IActiveScript::AddTypeLib");
            return E_NOTIMPL;
        }

        STDMETHOD(GetScriptDispatch)(LPCOLESTR pstrItemName, LPDISPATCH *ppdisp)
        {
            APSL_TRACE ("IActiveScript::GetCurrentScriptThreadID");
            if (!m_p_scriptsite_object)
                return E_POINTER;
            return m_p_scriptsite_object->get_member(pstrItemName, ppdisp);
        }

        STDMETHOD(GetCurrentScriptThreadID)(SCRIPTTHREADID *)
        {
            APSL_TRACE ("IActiveScript::GetCurrentScriptThreadID");
            return E_NOTIMPL;
        }

        STDMETHOD(GetScriptThreadID)(DWORD, SCRIPTTHREADID *)
        {
            APSL_TRACE ("IActiveScript::GetScriptThreadID");
            return E_NOTIMPL;
        }

        STDMETHOD(GetScriptThreadState)(SCRIPTTHREADID, SCRIPTTHREADSTATE *)
        {
            APSL_TRACE ("IActiveScript::GetScriptThreadState");
            return E_NOTIMPL;
        }

        STDMETHOD(InterruptScriptThread)(SCRIPTTHREADID, const EXCEPINFO *, DWORD)
        {
            APSL_TRACE ("IActiveScript::InterruptScriptThread");
            return E_NOTIMPL;
        }

        STDMETHOD(Clone)(IActiveScript **)
        {
            APSL_TRACE ("IActiveScript::Clone");
            return E_NOTIMPL;
        }

    public:
        aPSL::scriptsite_object *m_p_scriptsite_object;
        SCRIPTSTATE m_script_state;
        IActiveScriptSite *m_ActiveScriptSite;
        aPSL::script_engine *m_p_script_engine;

};

///////////////////////////////////////////////////////////////////////////
//
// @class IActiveScriptParseImpl
//
template <class T>
class __declspec(novtable) IActiveScriptParseImpl
: public IActiveScriptParse
{
public:
    IActiveScriptParseImpl()
    : m_p_activation_object(NULL)
    {
    }

    STDMETHOD(AddScriptlet)(
        LPCOLESTR pstrDefaultName,   // address of default name of scriptlet
        LPCOLESTR pstrCode,          // address of scriptlet text
        LPCOLESTR pstrItemName,      // address of item name
        LPCOLESTR pstrSubItemName,   // address of subitem name
        LPCOLESTR pstrEventName,     // address of event name
        LPCOLESTR pstrDelimiter,     // address of end-of-scriptlet delimiter
        DWORD dwSourceContextCookie, // application-defined value for debugging
        ULONG ulStartingLineNumber,  // starting line of the script
        DWORD dwFlags,               // scriptlet flags
        BSTR *pbstrName,             // address of actual name of scriptlet
        EXCEPINFO *pexcepinfo        // address of exception information
        )
    {
        HRESULT hr = S_OK;
        T* pthis = static_cast<T*>(this);
        if (0 == pthis->m_p_scriptsite_object)
            return E_POINTER;
        if (0 == pstrItemName || 0 == pstrEventName)
            return E_INVALIDARG;

        _bstr_t code = pstrItemName;
        if (pstrSubItemName) {
            return E_UNEXPECTED;
        }
        // TODO:
        //    code = code + L"." + pstrSubItemName;
        //code = code + L"." + pstrEventName + L"=function(){" + pstrCode + L"}";
        
        pthis->m_p_script_engine->eval(_bstr_t(code));
                
        return S_OK;
    }

    STDMETHOD(ParseScriptText)(
        LPCOLESTR pstrCode,
        LPCOLESTR pstrItemName,
        IUnknown *punkContext,
        LPCOLESTR pstrDelimiter,
        DWORD dwSourceContextCookie,
        ULONG ulStartingLineNumber,
        DWORD dwFlags,
        VARIANT *pvarResult,
        EXCEPINFO *pexcepinfo)
    {
        HRESULT hr = S_OK;
        T* pthis = static_cast<T*>(this);
        pthis->m_ActiveScriptSite->OnStateChange(
            pthis->m_script_state = SCRIPTSTATE_STARTED);
        pthis->m_p_script_engine->eval(_bstr_t(pstrCode));
        pthis->m_ActiveScriptSite->OnLeaveScript();
        pthis->m_script_state = SCRIPTSTATE_INITIALIZED;
	    pthis->m_ActiveScriptSite->OnStateChange(pthis->m_script_state);
        return S_OK;
    }

    STDMETHOD(InitNew)(VOID)
    {
        T *const pthis = static_cast<T*>(this);
        if (pthis->m_ActiveScriptSite)
            pthis->m_ActiveScriptSite->OnStateChange(
                pthis->m_script_state = SCRIPTSTATE_INITIALIZED);
        return S_OK;
    }

private:
    PSL::variable *m_p_activation_object;
};

////////////////////////////////////////////////////////////////////////////
//
// @class IActiveScriptGarbageCollectorImpl
//
template <class T>
class __declspec(novtable) IActiveScriptGarbageCollectorImpl
: public IActiveScriptGarbageCollector
{
// IActiveScriptGarbageCollector
        HRESULT STDMETHODCALLTYPE CollectGarbage(
            SCRIPTGCTYPE scriptgctype)
        {
            ::MessageBoxW(0,L"CollectGarbage",0,0);
            return S_OK;
        };
};

///////////////////////////////////////////////////////////////////////////
//
// @class CScriptObject
//
class __declspec(uuid(IID_APSL)) 
CScriptObject
    : public IActiveScriptImpl<CScriptObject>
    , public IActiveScriptParseImpl<CScriptObject>
    , public IActiveScriptGarbageCollectorImpl<CScriptObject>
{
public:
    INTERFACE_ENTRY const * GetInterfaceMap()
    {
        static INTERFACE_ENTRY interface_map[] = {
            { &__uuidof(IUnknown) , static_cast<IActiveScript *>(this) },
            { &__uuidof(IActiveScript) , static_cast<IActiveScript *>(this) },
            { &__uuidof(IActiveScriptParse) , static_cast<IActiveScriptParse *>(this) },
            { &__uuidof(IActiveScriptGarbageCollector) , static_cast<IActiveScriptGarbageCollector *>(this) },
            { NULL, NULL }
        };
        return interface_map;
    }
};

///////////////////////////////////////////////////////////////////////////
//
// @class CComFactory
//
template <typename T>
class CComFactory : public IClassFactory
{
public:
    CComFactory()
    : m_count(1)
    {
    }

    ~CComFactory() {}

//IUnknown
    STDMETHODIMP QueryInterface(REFIID riid, LPVOID *ppv)
    {
        if(IID_IUnknown != riid && IID_IClassFactory != riid)
            return E_NOINTERFACE;
        AddRef();
        *ppv = this;
        return S_OK;
    }

    STDMETHODIMP_(ULONG) AddRef(VOID)
    {
        return ++m_count;
    }

    STDMETHODIMP_(ULONG) Release(VOID)
    {
        if (--m_count == 0)
            delete this;
        return m_count;
    }

//IClassFactory
    STDMETHODIMP CreateInstance(
        IUnknown* pUnkOuter, REFIID riid, LPVOID *ppv)
    {
        if (pUnkOuter)
            return CLASS_E_NOAGGREGATION;
        T* p = new T;
        if(NULL == p)
            return E_OUTOFMEMORY;
        HRESULT hr = p->QueryInterface(riid, ppv);
        if (FAILED(hr))
            return hr;
        p->Release();
        return hr;
    }

    STDMETHODIMP LockServer(BOOL bLock)
    {
        return S_OK;
    }

private:
    LONG m_count;
};

const char *g_RegTable[][3] = {
  {"CLSID\\" IID_APSL, 0, PACKAGE_NAME " script language"},
  {"CLSID\\" IID_APSL "\\InProcServer32", 0, (const char*)-1},
  {"CLSID\\" IID_APSL "\\InProcServer32", "ThreadingModel", "Both"},
  {"CLSID\\" IID_APSL "\\OLEScript", 0, ""},
  {"CLSID\\" IID_APSL "\\ProgID", 0, PACKAGE_NAME},
  {PACKAGE_NAME, 0, PACKAGE_NAME " script language"},
  {PACKAGE_NAME "\\CLSID", 0, "{B7BCEFC5-FD47-4986-B418-A6686F9760CC}"},
  {PACKAGE_NAME "\\OLEScript", 0, ""},
};

static const GUID CATID_ActiveScript =
{ 0xf0b7a1a1, 0x9847, 0x11cf, { 0x8f, 0x20, 0x00, 0x80, 0x5f, 0x2c, 0xd0, 0x64 } };
static const GUID CATID_ActiveScriptParse =
{ 0xf0b7a1a2, 0x9847, 0x11cf, { 0x8f, 0x20, 0x00, 0x80, 0x5f, 0x2c, 0xd0, 0x64 } };

static const CATID ** GetCategoryMap() throw() 
{
    static const CATID * pMap[] = {
        &CATID_ActiveScript, &CATID_ActiveScriptParse, NULL
    };
    return pMap;
}

aPSL::com_module g_module;

//////////////////////////////////////////////////////////////////////////////
//
//  @fn     DllCanUnloadNow
//
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID )
{

    hInst = hInstance;
    if (dwReason == DLL_PROCESS_ATTACH)
        return 0 == g_module.increment();
    else if (dwReason == DLL_PROCESS_DETACH)
        return 0 == g_module.decrement();
    return TRUE;
}


//////////////////////////////////////////////////////////////////////////////
//
//  @fn     DllCanUnloadNow
//
STDAPI DllCanUnloadNow(VOID)
{
    return g_module.can_unload_now() == 0 ? S_OK : S_FALSE;
}

//////////////////////////////////////////////////////////////////////////////
//
//  @fn     DllGetClassObject
//
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
    CComFactory<IUnknownImpl<CScriptObject> > *pCF
        = new CComFactory<IUnknownImpl<CScriptObject> >();
    if (pCF != NULL)
        return pCF->QueryInterface(riid, ppv);
    return S_OK;
}

//////////////////////////////////////////////////////////////////////////////
//
//  @fn     DllUnregisterServer
//
STDAPI DllUnregisterServer(VOID)
{
    HRESULT hr = S_FALSE;
    ICatRegister *pCatRegister;
    hr = CoCreateInstance(
        CLSID_StdComponentCategoriesMgr, NULL,
	    CLSCTX_INPROC_SERVER, __uuidof(ICatRegister),
	    reinterpret_cast<LPVOID*>(&pCatRegister) );
    if (FAILED(hr) || !pCatRegister)
	    return E_FAIL;
    pCatRegister->AddRef();
    for (CATID const** pEntry = GetCategoryMap(); *pEntry; ++pEntry)
        if (FAILED(hr = pCatRegister->UnRegisterClassImplCategories(
            __uuidof(CScriptObject), 1, const_cast<CATID *>(*pEntry))))
            break;
    pCatRegister->Release();
    if (FAILED(hr))
	    return hr;
    for (int i = sizeof(g_RegTable)/sizeof(*g_RegTable) - 1; i >= 0; --i)
        if(RegDeleteKeyA(HKEY_CLASSES_ROOT, g_RegTable[i][0]) != ERROR_SUCCESS)
            hr = S_FALSE;
    return S_OK;
}

//////////////////////////////////////////////////////////////////////////////
//
//  @fn     DllRegisterServer
//
STDAPI DllRegisterServer(VOID)
{
    HRESULT hr = S_FALSE;
    ICatRegister *pCatRegister;
    hr = CoCreateInstance(
        CLSID_StdComponentCategoriesMgr, NULL,
	    CLSCTX_INPROC_SERVER,
	    __uuidof(ICatRegister),
	    reinterpret_cast<void **>(&pCatRegister) );
    if (FAILED(hr))
	    return hr;
    pCatRegister->AddRef();
    for (CATID const** pEntry = GetCategoryMap(); *pEntry; ++pEntry)
        if (FAILED(hr = pCatRegister->RegisterClassImplCategories(
            __uuidof(CScriptObject), 1, const_cast<CATID *>(*pEntry))))
            break;
    pCatRegister->Release();
    if (FAILED(hr))
	    return hr;
    CHAR szFileName[MAX_PATH] = { 0 };
    GetModuleFileNameA(hInst, szFileName, MAX_PATH);
    for (int i = 0; SUCCEEDED(hr) && i < sizeof(g_RegTable) / sizeof(*g_RegTable); ++i)
    {
        LPCSTR pszKeyName   = g_RegTable[i][0];
        LPCSTR pszValueName = g_RegTable[i][1];
        LPCSTR pszValue     = g_RegTable[i][2];
        if(pszValue == (LPCSTR)-1)
            pszValue = szFileName;
        HKEY hkey;
        if (RegCreateKeyA(HKEY_CLASSES_ROOT, pszKeyName, &hkey)
            != ERROR_SUCCESS)
            return DllUnregisterServer(), SELFREG_E_CLASS;
        RegSetValueExA(hkey, pszValueName, 0, REG_SZ, (const BYTE*)pszValue,
            (strlen(pszValue) + 1));
        RegCloseKey(hkey);
    }
    return S_OK;
}


