#include "CXl.h"
#include <comdef.h> // _bstr_t
#include <regex>    // regex, smatch
#include <vector>   // vector

using namespace std;

unsigned int* ConvertXlNumber(const char* psz) {
    vector<string> vsBuf;
    string strColumn;
    int nBuf = 0;
    unsigned int unRow = 0;
    smatch match;

    for (int i = 0; i < strlen(psz); i++) {
        vsBuf.emplace_back(string(1, psz[i]));

        if (regex_match(vsBuf[i], match, regex("[A-Z]")))
            nBuf++;
        else if (regex_match(vsBuf[i], match, regex("[0-9]")))
            strColumn += psz[i];
    }

    unsigned int power[3] = { 1,26,26 * 26 };

    for (int i = 0, j = nBuf - 1; i < nBuf; i++, j--)
        unRow += (psz[i] - 64) * power[j];

    unsigned int result[2] = { unRow, stoi(strColumn) };

    return result;
}

// Able_CommonFunc.cpp charTowchar()
wchar_t* WIDE(const char* psz) {
    wchar_t* pwsz;
    int n = MultiByteToWideChar(CP_ACP, 0, psz, -1, NULL, 0);
    pwsz = new wchar_t[n];
    MultiByteToWideChar(CP_ACP, 0, psz, -1, pwsz, n);
    /*
    size_t si = strlen(psz) + 1;
    wchar_t* pwsz = (LPWSTR)malloc(sizeof(wchar_t) * si);
    mbstowcs_s(NULL, pwsz, si, psz, si-1);
    */
    return pwsz;
}

// Able_CommonFunc.cpp formatstring()
// docs.microsoft.com FormatOutput()
const char* FormatOutput(const char* psz, ...) {
    char buf[4096];

    va_list vargs;
    va_start(vargs, psz);
    vsnprintf_s(buf, _TRUNCATE, psz, vargs);
    va_end(vargs);

    return buf;
}

// 오류 코드 모르니 무슨 뜻인지도 모릅니다. 수정이 필요합니다
bool CheckError(HRESULT hr) {
    if (SUCCEEDED((DWORD)hr))
        return false;

    SetLastError((DWORD)hr);

    LPVOID lpMsgBuf;
    FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER |
        FORMAT_MESSAGE_FROM_SYSTEM |
        FORMAT_MESSAGE_IGNORE_INSERTS,
        NULL,
        (DWORD)hr,
        0,
        (LPWSTR)&lpMsgBuf,
        0,
        NULL);

    MessageBox(NULL, (LPCTSTR)lpMsgBuf, L"오류", MB_OK | MB_ICONERROR);

    LocalFree(lpMsgBuf);

    return true;
}

HRESULT AutoWrap(int autoType, VARIANT* pvResult, IDispatch* pDisp, const char* pszName, int cArgs...) {
    // LPOLESTR로 일일히 형변환하기 싫었기 때문에 추가한 코드입니다.
    LPOLESTR ptName = WIDE(pszName);

    // Begin variable-argument list...
    va_list marker;
    va_start(marker, cArgs);

    if (!pDisp) {
        MessageBoxA(NULL, "AutoWrap()의 IDispatch 인자가 존재하지 않습니다.", "오류", 0x10010);
        return E_INVALIDARG;
    }

    // Variables used...
    DISPPARAMS dp = { NULL, NULL, 0, 0 };
    DISPID dispidNamed = DISPID_PROPERTYPUT;
    DISPID dispID;
    HRESULT hr;

    // Get DISPID for name passed...
    if (CheckError(hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID)))
        return hr;

    // Allocate memory for arguments...
    VARIANT* pArgs = new VARIANT[(size_t)cArgs + 1];
    // Extract arguments...
    for (int i = 0; i < cArgs; i++) {
        pArgs[i] = va_arg(marker, VARIANT);
    }

    // Build DISPPARAMS
    dp.cArgs = cArgs;
    dp.rgvarg = pArgs;

    // Handle special-case for property-puts!
    if (autoType & DISPATCH_PROPERTYPUT) {
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
    }

    // Make the call!
    if (CheckError(hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, NULL, NULL)))
        return hr;

    // End variable-argument section...
    va_end(marker);

    delete[] pArgs;

    return hr;
}

CXl::CXl() {
    CheckError(m_hrCOMInit = CoInitialize(NULL));
}

CXl::~CXl() {
    if (m_hrCOMInit == S_OK)
        CoUninitialize();
}

int CXl::Init() {
    CLSID clsid;

    if (CheckError(CLSIDFromProgID(L"Excel.Application", &clsid)))
        return -1;
    if (CheckError(CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&m_XlProp.pXlApp)))
        return -2;

    VARIANT result;
    VariantInit(&result);
    if (CheckError(m_hrInit = AutoWrap(DISPATCH_PROPERTYGET, &result, m_XlProp.pXlApp, "Workbooks", 0)))
        return -3;

    m_XlProp.pXlBooks = result.pdispVal;

    return 1;
}

int CXl::SetVisible(bool bVisible) {
    if (CheckError(CheckXlInit()))
        return -1;

    VARIANT x;
    x.vt = VT_I4;
    x.lVal = bVisible ? 1 : 0;
    if (CheckError(AutoWrap(DISPATCH_PROPERTYPUT, NULL, m_XlProp.pXlApp, "Visible", 1, x)))
        return -2;

    return 1;
}

int CXl::AddWorkBooks() {
    VARIANT result;
    VariantInit(&result);
    if (CheckError(AutoWrap(DISPATCH_PROPERTYGET, &result, m_XlProp.pXlBooks, "Add", 0)))
        return -1;

    m_XlProp.pXlBook = result.pdispVal;

    return 1;
}

int CXl::Open()
{
    if (Init() < 0)
        return -10;
    if (CheckError(CheckXlInit()))
        return -20;

    if (SetVisible(true) < 0)
        return -30;

    if (AddWorkBooks() < 0)
        return -40;

    VariantInit(&m_XlProp.pvaArray);
    m_XlProp.pvaArray.vt = VT_ARRAY | VT_VARIANT;

    return 1;
}

int CXl::SetSafeBound(const char* pcszStart, const char* pcszEnd) {

    unsigned int unRowMin = ConvertXlNumber(pcszStart)[0];
    unsigned int unRowMax = ConvertXlNumber(pcszEnd)[0];
    unsigned int unColumnMin = ConvertXlNumber(pcszStart)[1];
    unsigned int unColumnMax = ConvertXlNumber(pcszEnd)[1];
    return SetSafeBound(unRowMin, unRowMax, unColumnMin, unColumnMax);
}

int CXl::SetSafeBound(unsigned int unRowMin, unsigned int unRowMax, unsigned int unColumnMin, unsigned int unColumnMax) {
    SAFEARRAYBOUND sab[2];
    printf("unRowMin: %d\n", unRowMin);
    printf("unRowMax: %d\n", unRowMax);
    printf("unColumnMin: %d\n", unColumnMin);
    printf("unColumnMax: %d\n", unColumnMax);

    sab[0].lLbound = unRowMin; sab[0].cElements = unRowMax;
    sab[1].lLbound = unColumnMin; sab[1].cElements = unColumnMax;

    if ((m_XlProp.pvaArray.parray = SafeArrayCreate(VT_VARIANT, 2, sab)) == NULL) {
        MessageBoxA(NULL, "SafeArrayCreat()의 값을 반환하지 못했습니다.", "오류", 0x10010);
        return -1;
    }

    return 1;
}

int CXl::SetRange(const char* pcszStart, const char* pcszEnd) {
    _bstr_t str = FormatOutput("%s:%s", pcszStart, pcszEnd);
    VARIANT range;
    range.vt = VT_BSTR;
    range.bstrVal = ::SysAllocString(str);

    VARIANT result;
    VariantInit(&result);
    if (CheckError(AutoWrap(DISPATCH_PROPERTYGET, &result, m_XlProp.pXlSheet, "Range", 1, range)))
        return -3;

    VariantClear(&range);

    m_XlProp.pXlRange = result.pdispVal;

    return 1;
}

int CXl::AddActiveSheet() {
    VARIANT result;
    VariantInit(&result);
    if (CheckError(AutoWrap(DISPATCH_PROPERTYGET, &result, m_XlProp.pXlApp, "ActiveSheet", 0)))
        return -1;

    m_XlProp.pXlSheet = result.pdispVal;

    return 1;
}

int CXl::SetData(const char* pcszData, long lRow, long lColumn) {
    _bstr_t bstrData = pcszData;
    VARIANT vaData;
    vaData.vt = VT_BSTR;
    vaData.bstrVal = ::SysAllocString(bstrData);

    long indices[] = { lRow, lColumn };

    if (CheckError(SafeArrayPutElement(m_XlProp.pvaArray.parray, indices, (void*)&vaData)))
        return -1;

    return 1;
}

int CXl::Print() {
    if (CheckError(AutoWrap(DISPATCH_PROPERTYPUT, NULL, m_XlProp.pXlRange, "Value", 1, m_XlProp.pvaArray)))
        return -1;

    return 1;
}

int CXl::Save() {
    if (CheckError(AutoWrap(DISPATCH_METHOD, NULL, m_XlProp.pXlApp, "Save", 0)))
        return -1;

    return 1;
}

int CXl::Close() {
    if (CheckError(AutoWrap(DISPATCH_METHOD, NULL, m_XlProp.pXlApp, "Quit", 0)))
        return -1;

    m_XlProp.pXlRange->Release();
    m_XlProp.pXlSheet->Release();
    m_XlProp.pXlBook->Release();
    m_XlProp.pXlBooks->Release();
    m_XlProp.pXlApp->Release();
    VariantClear(&m_XlProp.pvaArray);

    Sleep(3000);

    CoUninitialize();
    MessageBoxA(NULL, "성공적으로 작업이 종료되었습니다", "알림", MB_OK | MB_ICONINFORMATION);

    return 0;
}