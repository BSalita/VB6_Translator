#include "windows.h"
#include "wchar.h"

/* Beware: These routines are partial implementations */

void e(char *s) {printf("undefs: not implemented: %s\n",s); exit(666);}
#define def(n,s) void n(){e(s);}
def(_findclose,"_findclose")
def(_wrmdir,"_wrmdir")
BOOL SetSystemTime(CONST SYSTEMTIME *lpSystemTime)
{
	e("SetSystemTime");
}
def(_wmkdir,"_wmkdir")
def(_wstat,"_wstat")
def(_wgetdcwd,"_wgetdcwd")
def(_wfindfirst,"_wfindfirst")
def(VarFormatNumber,"VarFormatNumber")
def(VarFormatCurrency,"VarFormatCurrency")
def(VarMonthName,"VarMonthName")
def(VarAbs,"VarAbs")
wchar_t *_wgetcwd(wchar_t *wbuf,int maxlen)
{
char buf[1024]; /* use named const instead of 1024? */
	if (getcwd(buf,1024) == NULL)
		return(NULL);
	mbstowcs(wbuf,buf,maxlen);
vbt_printf("_wgetcwd: cwd=%S\n",wbuf);
	return wbuf;
}
def(_wremove,"_wremove")

int _wchdir(const wchar_t *dirname)
{
char buf[1024]; /* use named const instead of 1024? */
vbt_printf("_wchdir: dirname=%S\n",dirname);
	wcstombs(buf,dirname,sizeof(buf));
	if (chdir(buf))
		return(-1);
vbt_printf("_wchdir: cwd=%s\n",getcwd(buf,sizeof(buf)));
	return(0);
}

HRESULT VarInt(LPVARIANT pvarIn, LPVARIANT  pvarResult)
{
VARIANT v;
HRESULT hr;
DOUBLE floor(DOUBLE d);
	VariantInit(&v);
	hr = VariantChangeType(&v,pvarIn,0,VT_R8);
	if (SUCCEEDED(hr))
		{
		V_R8(&v) = floor(V_R8(&v));
		hr = VariantChangeType(pvarResult,&v,0,V_VT(pvarIn));
		}
	return(hr);
}
 
wchar_t *_wfullpath( wchar_t *absPath, const wchar_t *relPath, size_t maxLength )
{
vbt_printf("_wfullpath: absPath=%S relPath=%S maxLength=%u\n",absPath,relPath,maxLength);
	if (*relPath == L'/')
		wcsncpy(absPath,relPath,maxLength);
	else
	{
		if (_wgetcwd(absPath,maxLength) == NULL)
			return NULL;
		wcsncat(absPath,L"/",maxLength);
		wcsncat(absPath,relPath,maxLength);
	}
	while (absPath[wcslen(absPath+1)] == L'/') 
		absPath[wcslen(absPath+1)] = 0;
vbt_printf("absPath=%S\n",absPath);
	return absPath;
}

def(VarFormat,"VarFormat")
def(_wfindnext,"_wfindnext")

HRESULT VarBstrCat(BSTR bstrLeft, BSTR bstrRight, LPBSTR pbstrResult)
{
	*pbstrResult = SysAllocStringLen(bstrLeft,SysStringLen(bstrLeft)+SysStringLen(bstrRight));
	wcsncpy(*pbstrResult+SysStringLen(bstrLeft),bstrRight,SysStringLen(bstrRight));
	return(0);
}

def(VarFormatPercent,"VarFormatPercent")
def(ShellExecute,"ShellExecute")
def(VarFormatDateTime,"VarFormatDateTime")
def(VarFix,"VarFix")
def(VarWeekdayName,"VarWeekdayName")
def(_wgetenv,"_wgetenv")
def(_chdrive,"_chdrive")

wchar_t * _itow( int value, wchar_t *string, int radix )
{
wchar_t *w = string;
vbt_printf("_itow=%d string=%lx radix=%d\n",value,string,radix);
	do
	{
		*w++ = L"01234567890ABCDEF"[value % radix];
		value /= radix;
	} while(value);
	*w = 0;
vbt_printf("w=%S\n",string);
	return(string);
}

def(VarRound,"VarRound")
