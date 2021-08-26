
/*
to do:
1. implement Null values
2. remove static values
3. add error checking
4. should date/time funcs use ANSI C or Win32 APIs?
5. implement ERR_NOT_IMPLMENTED items
6. CreateInstance s/b CreateInstanceEx
*/

#define INCLUDE_INTERFACES
#include "test3/c/vba_i.h" /* where to put vba.c and vba_i*??? */
#include <time.h> /* clock */
#ifndef SAG_COM
#include <math.h> /* log */
#include <direct.h> /* _chdrive */
#include <wchar.h> /* _wstat */
#include <io.h> /* _findclose */
#include <errno.h> /* EEXIST */
#include <shellapi.h> /* ShellExecute */
#else
#include <errno.h> /* EEXIST */
#include <sys/stat.h> /* struct stat */
#define _stat stat
struct _wfinddata_t
{
	wchar_t name[260];
	unsigned attrib;
};
int *_wchdir(const wchar_t *dirname);
long _wfindfirst(wchar_t *filespec, struct _wfinddata_t *fileinfo);
long _wfindnext(long handle, struct _wfinddata_t *fileinfo);
wchar_t *_wgetcwd(wchar_t *buffer, int maxlen);
wchar_t *_wgetdcwd(int drive, wchar_t *buffer, int maxlen);
wchar_t *_wgetenv(const wchar_t *varname);
wchar_t _wmkdir(const wchar_t *dirname);
wchar_t *_wremove(const wchar_t *path);
wchar_t _wstat(const wchar_t *path, struct _stat *buffer);
#endif

wchar_t *wmemmove(wchar_t *s1, const wchar_t *s2, size_t n);
wchar_t *wmemset(wchar_t *s, wchar_t c, size_t n);
typedef short INT16;
#define diffptr(a,b) (size_t)((CHAR *)(a)-(CHAR *)(b))
#define SW_SECS_PER_DAY (24*60*60)
#define VB_SIGNIFICANCE 15
int vbt_diag(int errnum,char *file,int line);
#define diag(errnum) vbt_diag(0/*errnum*/,__FILE__,__LINE__)
INT vbt_printf(const char *format, ...);
struct tm *vbdatetotm(DOUBLE vbdate);
double tmtovbdate(struct tm *tm);
extern int g_argc;
extern wchar_t *g_wargv[];
extern wchar_t *g_wenvp[];
/* define in vbio.c */
#define FC short
struct filblk_t { int fc; };
typedef struct filblk_t FILBLK;
FILBLK *iofclu(FC fc){diag(ERR_NOT_IMPLEMENTED);}
long iofileattr(FC fc,int attr){diag(ERR_NOT_IMPLEMENTED);}
INT ioclsall(VOID){diag(ERR_NOT_IMPLEMENTED);}
VOID iocopy(CHAR *sfrom, size_t lfrom, CHAR *sto, size_t lto){diag(ERR_NOT_IMPLEMENTED);}
INT ioeof(FC fc){diag(ERR_NOT_IMPLEMENTED);}
void fb_tell(FILBLK *fb, long *off){diag(ERR_NOT_IMPLEMENTED);}
void fb_stat(FILBLK *fb, struct stat *statbuf){diag(ERR_NOT_IMPLEMENTED);}

/* Declaration: Strings  0 . 0  */
short WINAPI VBA_Asc(BSTR String)
{
	return(String[0]);
}

BSTR WINAPI VBA__B_str_Chr(long CharCode)
{
BSTR bstr;
	VarBstrFromI4(CharCode,0,0,&bstr);
	return bstr;
}

VARIANT WINAPI VBA__B_var_Chr(long CharCode)
{
VARIANT v;
	VariantInit(&v);
	V_VT(&v) = VT_BSTR;
	VarBstrFromI4(CharCode,0,0,&V_BSTR(&v));
	return v;
}

BSTR WINAPI VBA__B_str_LCase(BSTR String)
{
BSTR bstr;
	bstr = SysAllocString(String);
	wcslwr(bstr);
	return bstr;
}

VARIANT WINAPI VBA__B_var_LCase(VARIANT* String)
{
VARIANT v;
	VariantChangeType(&v,String,0,VT_BSTR);
	wcslwr(V_BSTR(&v));
	return v;
}

BSTR WINAPI VBA__B_str_Mid(BSTR String,long Start,VARIANT* Length)
{
size_t len,slen;
	slen = SysStringLen(String);
	if (Start < 0)
		abort(); /* invalid function value */
	if (V_VT(Length) == VT_EMPTY)
		{
		len = slen;
		}
	else
		{
		VARIANT vLength;
		VariantInit(&vLength);
		VariantChangeType(&vLength,Length,0,VT_I4);
		len = V_I4(&vLength);
/*		VariantClear(&vLength);*/ /* not needed for VT_I4 */
		}
	if ((size_t)Start > slen)
		len = 0;
	else if (Start-1+len > slen)
		len = slen-(Start-1);
	return SysAllocStringLen(String+Start-1,len);
}

VARIANT WINAPI VBA__B_var_Mid(VARIANT* String,long Start,VARIANT* Length)
{
VARIANT vString,vresult;
size_t len,slen;
	VariantInit(&vString);
	VariantChangeType(&vString,String,0,VT_BSTR);
	slen = SysStringLen(V_BSTR(&vString));
	if (Start < 0)
		abort(); /* invalid function value */
	if (V_VT(Length) == VT_EMPTY)
		{
		len = slen;
		}
	else
		{
		VARIANT vLength;
		VariantInit(&vLength);
		VariantChangeType(&vLength,Length,0,VT_I4);
		len = V_I4(&vLength);
/*		VariantClear(&vLength);*/ /* not needed for VT_I4 */
		}
	if ((size_t)Start > slen)
		len = 0;
	else if (Start-1+len > slen)
		len = slen-(Start-1);
/*	VariantInit(&vresult);*/ /* not needed */
	V_VT(&vresult) = VT_BSTR;
	V_BSTR(&vresult) = SysAllocStringLen(V_BSTR(&vString)+Start-1,len);
	VariantClear(&vString);
	return vresult;
}

BSTR WINAPI VBA__B_str_MidB(BSTR String,long Start,VARIANT* Length)
{
size_t len,slen;
	slen = SysStringByteLen(String);
	if (Start < 0)
		abort(); /* invalid function value */
	if (V_VT(Length) == VT_EMPTY)
		{
		len = slen;
		}
	else
		{
		VARIANT vLength;
		VariantInit(&vLength);
		VariantChangeType(&vLength,Length,0,VT_I4);
		len = V_I4(&vLength);
/*		VariantClear(&vLength);*/ /* not needed for VT_I4 */
		}
	if ((size_t)Start > slen)
		len = 0;
	else if (Start-1+len > slen)
		len = slen-(Start-1);
	return SysAllocStringByteLen((CHAR *)String+Start-1,len);
}

VARIANT WINAPI VBA__B_var_MidB(VARIANT* String,long Start,VARIANT* Length)
{
VARIANT vString,vresult;
size_t len,slen;
	VariantInit(&vString);
	VariantChangeType(&vString,String,0,VT_BSTR);
	slen = SysStringByteLen(V_BSTR(&vString));
	if (Start < 0)
		abort(); /* invalid function value */
	if (V_VT(Length) == VT_EMPTY)
		{
		len = slen;
		}
	else
		{
		VARIANT vLength;
		VariantInit(&vLength);
		VariantChangeType(&vLength,Length,0,VT_I4);
		len = V_I4(&vLength);
/*		VariantClear(&vLength);*/ /* not needed for VT_I4 */
		}
	if ((size_t)Start > slen)
		len = 0;
	else if (Start-1+len > slen)
		len = slen-(Start-1);
/*	VariantInit(&vresult);*/ /* not needed */
	V_VT(&vresult) = VT_BSTR;
	V_BSTR(&vresult) = SysAllocStringByteLen((CHAR *)V_BSTR(&vString)+Start-1,len);
	VariantClear(&vString);
	return vresult;
}

VARIANT WINAPI VBA_InStr(VARIANT* Start,VARIANT* String1,VARIANT* String2,VBA_VbCompareMethod Compare)
{
VARIANT vStart, vString1, vString2, vResult;

	VariantChangeType(&vStart, Start, 0, VT_I4);
	VariantChangeType(&vString1, String1, 0, VT_BSTR);
	VariantChangeType(&vString2, String2, 0, VT_BSTR);
	V_VT(&vResult) = VT_I4;
	if (V_I4(&vStart) <= 0)
		diag(ERR_INVALID_FUNCTION_VALUE);
	if (*V_BSTR(&vString2) == 0)
		V_I4(&vResult) = V_I4(&vStart);
	else if ((size_t)V_I4(&vStart) > wcslen(V_BSTR(&vString1)))
		V_I4(&vResult) = 0;
	else
		{
		WCHAR *p;
		if (Compare)
			diag(ERR_NOT_IMPLEMENTED);
		else
			p = wcswcs(V_BSTR(&vString1)+V_I4(&vStart)-1,V_BSTR(&vString2)); /* wcsstr */
		if (p == NULL)
			V_I4(&vResult) = 0;
		else
			V_I4(&vResult) = p-V_BSTR(&vString1)+1;
		}
	return vResult;
}

VARIANT WINAPI VBA_InStrB(VARIANT* Start,VARIANT* String1,VARIANT* String2,VBA_VbCompareMethod Compare)
{
VARIANT vStart, vString1, vString2, vResult;

	VariantChangeType(&vStart, Start, 0, VT_I4);
	VariantChangeType(&vString1, String1, 0, VT_BSTR);
	VariantChangeType(&vString2, String2, 0, VT_BSTR);
	V_VT(&vResult) = VT_I4;
	if (V_I4(&vStart) <= 0)
		diag(ERR_INVALID_FUNCTION_VALUE);
	if (*V_BSTR(&vString2) == 0)
		V_I4(&vResult) = V_I4(&vStart);
	else if ((size_t)V_I4(&vStart) > strlen((CHAR *)V_BSTR(&vString1)))
		V_I4(&vResult) = 0;
	else
		{
		CHAR *p;
		if (Compare)
			diag(ERR_NOT_IMPLEMENTED);
		else
			p = strstr((CHAR *)V_BSTR(&vString1)+V_I4(&vStart)-1,(CHAR *)V_BSTR(&vString2)); /* wcsstr */
		if (p == NULL)
			V_I4(&vResult) = 0;
		else
			V_I4(&vResult) = p-(CHAR *)V_BSTR(&vString1)+1;
		}
	return vResult;
}

BSTR WINAPI VBA__B_str_Left(BSTR String,long Length)
{
size_t slen;

	if (Length < 0)
		diag(ERR_INVALID_FUNCTION_VALUE);
	slen = SysStringLen(String);
	return SysAllocStringLen(String, (size_t)Length <= slen ? Length : slen);
}

VARIANT WINAPI VBA__B_var_Left(VARIANT* String,long Length)
{
size_t slen;
VARIANT vString,vResult;

	if (Length < 0)
		diag(ERR_INVALID_FUNCTION_VALUE);
	VariantChangeType(&vString,String,0,VT_BSTR);
	slen = SysStringLen(V_BSTR(&vString));
	V_VT(&vResult) = VT_BSTR;
	V_BSTR(&vResult) = SysAllocStringLen(V_BSTR(&vString), (size_t)Length <= slen ? Length : slen);
	return vResult;
}

BSTR WINAPI VBA__B_str_LeftB(BSTR String,long Length)
{
size_t slen;

	if (Length < 0)
		diag(ERR_INVALID_FUNCTION_VALUE);
	slen = SysStringByteLen(String);
	return SysAllocStringByteLen((CHAR *)String, (size_t)Length <= slen ? Length : slen);
}

VARIANT WINAPI VBA__B_var_LeftB(VARIANT* String,long Length)
{
size_t slen;
VARIANT vString, vResult;

	if (Length < 0)
		diag(ERR_INVALID_FUNCTION_VALUE);
	VariantChangeType(&vString,String,0,VT_BSTR);
	slen = SysStringByteLen(V_BSTR(&vString));
	V_VT(&vResult) = VT_BSTR;
	V_BSTR(&vResult) = SysAllocStringByteLen((CHAR *)V_BSTR(&vString), (size_t)Length <= slen ? Length : slen);
	return vResult;
}

BSTR WINAPI VBA__B_str_LTrim(BSTR String)
{
	return SysAllocString(String+wcsspn(String,L" "));
}

VARIANT WINAPI VBA__B_var_LTrim(VARIANT* String)
{
VARIANT vString,vResult;
	VariantChangeType(&vString,String,0,VT_BSTR);
	V_VT(&vResult) = VT_BSTR;
	V_BSTR(&vResult) = SysAllocString(V_BSTR(&vString)+wcsspn(V_BSTR(&vString),L" "));
	return vResult;
}

BSTR WINAPI VBA__B_str_RightB(BSTR String,long Length)
{
size_t slen;

	if (Length < 0)
		diag(ERR_INVALID_FUNCTION_VALUE);
	slen = SysStringByteLen(String);
	return SysAllocStringByteLen((CHAR *)String + ((size_t)Length <= slen ? slen-Length : slen),(size_t)Length <= slen ? Length : slen);
}

VARIANT WINAPI VBA__B_var_RightB(VARIANT* String,long Length)
{
VARIANT vString, vResult;
size_t slen;

	if (Length < 0)
		diag(ERR_INVALID_FUNCTION_VALUE);
	VariantChangeType(&vString,String,0,VT_BSTR);
	slen = SysStringByteLen(V_BSTR(&vString));
	V_VT(&vResult) = VT_BSTR;
	V_BSTR(&vResult) = SysAllocStringByteLen((CHAR *)V_BSTR(&vString) + ((size_t)Length <= slen ? slen-Length : slen),(size_t)Length <= slen ? Length : slen);
	return vResult;
}

BSTR WINAPI VBA__B_str_Right(BSTR String,long Length)
{
size_t slen;

	if (Length < 0)
		diag(ERR_INVALID_FUNCTION_VALUE);
	slen = SysStringLen(String);
	return SysAllocString(String + ((size_t)Length <= slen ? slen-Length : slen));
}

VARIANT WINAPI VBA__B_var_Right(VARIANT* String,long Length)
{
VARIANT vString, vResult;
size_t slen;

	if (Length < 0)
		diag(ERR_INVALID_FUNCTION_VALUE);
	VariantChangeType(&vString,String,0,VT_BSTR);
	slen = SysStringLen(V_BSTR(&vString));
	V_VT(&vResult) = VT_BSTR;
	V_BSTR(&vResult) = SysAllocString(V_BSTR(&vString) + ((size_t)Length <= slen ? slen-Length : slen));
	return vResult;
}

BSTR WINAPI VBA__B_str_RTrim(BSTR String)
{
OLECHAR *w;
	w = String+SysStringLen(String);
	while(w != String)
		if (*--w != L' ')
			break;
	return SysAllocStringLen(String,diffptr(w,String));
}

VARIANT WINAPI VBA__B_var_RTrim(VARIANT* String)
{
VARIANT vString,vResult;
OLECHAR *w;
	VariantChangeType(&vString,String,0,VT_BSTR);
	V_VT(&vResult) = VT_BSTR;
	w = V_BSTR(&vString)+SysStringLen(V_BSTR(&vString));
	while(w != V_BSTR(&vString))
		if (*--w != L' ')
			break;
	V_BSTR(&vResult) = SysAllocStringLen(V_BSTR(&vString),diffptr(w,V_BSTR(&vString)));
	return vResult;
}

BSTR WINAPI VBA__B_str_Space(long Number)
{
BSTR bstr;
	bstr = SysAllocStringLen(NULL,Number);
	wmemset(bstr,L' ',Number);
	return bstr;
}

VARIANT WINAPI VBA__B_var_Space(long Number)
{
VARIANT vResult;
	V_VT(&vResult) = VT_BSTR;
	V_BSTR(&vResult) = SysAllocStringLen(NULL,Number);
	wmemset(V_BSTR(&vResult),L' ',Number);
	return vResult;
}

VARIANT WINAPI VBA__B_var_StrConv(VARIANT* String,VBA_VbStrConv Conversion,long LocaleID)
{
VARIANT vResult,vString;
BSTR bstr,newbstr;
size_t len;
CHAR *buf;
	VariantChangeType(&vString,String,0,VT_BSTR);
	bstr = V_BSTR(&vString);
	switch(Conversion)
	{
	case 1: /* vbUpperCase */
		wcsupr(bstr);
		break;
	case 2: /* vbLowerCase */
		wcslwr(bstr);
		break;
	case 3: /* vbProperCase */
		{
		BSTR s = bstr;
		BSTR se = bstr+SysStringLen(bstr);
		while(s<se)
			{
			while(s<se && (iswspace(*s) || *s == 0))
				s++;
			if (s < se)
				*s = towupper(*s); /* doesn't have to be upperable */
			while(s<se && !(iswspace(*s) || *s == 0))
				s++;
			}
		}
		break;
	case 4: /* vbWide */
	case 8: /* vbNarrow */
	case 16: /* vbKatakana */
	case 32: /* vbHiragana */
		VariantClear(&vString);
		diag(ERR_NOT_IMPLEMENTED);
		break;
	case 64: /* vbUnicode */
		len = SysStringLen(bstr)*sizeof(WCHAR);
		newbstr = SysAllocStringLen(NULL,len);
		newbstr[MultiByteToWideChar(CP_ACP,MB_PRECOMPOSED,(CHAR *)bstr,len,newbstr,len)] = 0;
		bstr = newbstr;
		break;
	case 128: /* vbFromUnicode */
		len = SysStringLen(bstr);
		buf = (CHAR *)SysAllocStringLen(NULL,(len+1)/sizeof(WCHAR)); /* use SysAllocStringByteLen instead */
		buf[WideCharToMultiByte(CP_ACP,0,bstr,len,buf,len,NULL,NULL)] = 0;
		bstr = (BSTR)buf;
		bstr[(len+1)/sizeof(WCHAR)] = 0;
		break;
	default:
		VariantClear(&vString);
		diag(ERR_SYSTEM_ERROR);
	}
	V_VT(&vResult) = VT_BSTR;
	V_BSTR(&vResult) = bstr;
	VariantClear(&vString);
	return vResult;
}

BSTR WINAPI VBA__B_str_String(long Number,VARIANT* Character)
{
VARIANT vString;
BSTR bstr;

	VariantChangeType(&vString,Character,0,VT_BSTR);
	bstr = SysAllocStringLen(NULL,Number);
	wmemset(bstr,*V_BSTR(&vString),Number);
	VariantClear(&vString);
	bstr[Number] = 0;
	return bstr;
}

VARIANT WINAPI VBA__B_var_String(long Number,VARIANT* Character)
{
VARIANT vResult,vString;

	VariantChangeType(&vString,Character,0,VT_BSTR);
	V_VT(&vResult) = VT_BSTR;
	V_BSTR(&vResult) = SysAllocStringLen(NULL,Number);
	wmemset(V_BSTR(&vResult),*V_BSTR(&vString),Number);
	VariantClear(&vString);
	V_BSTR(&vResult)[Number] = 0;
	return vResult;
}

BSTR WINAPI VBA__B_str_Trim(BSTR String)
{
OLECHAR *l,*r;
	l = String+wcsspn(String,L" ");
	r = l+SysStringLen(l);
	while(r != l)
		if (*--r != L' ')
			break;
	return SysAllocStringLen(l,diffptr(r,l));
}

VARIANT WINAPI VBA__B_var_Trim(VARIANT* String)
{
OLECHAR *l,*r;
VARIANT vString,vResult;

	VariantChangeType(&vString,String,0,VT_BSTR);
	l = V_BSTR(&vString)+wcsspn(V_BSTR(&vString),L" ");
	r = l+SysStringLen(l);
	while(r != l)
		if (*--r != L' ')
			break;
	V_VT(&vResult) = VT_BSTR;
	V_BSTR(&vResult) = SysAllocStringLen(l,diffptr(r,l));
	VariantClear(&vString);
	return vResult;
}

BSTR WINAPI VBA__B_str_UCase(BSTR String)
{
BSTR bstr;
	bstr = SysAllocString(String);
	wcsupr(bstr);
	return bstr;
}

VARIANT WINAPI VBA__B_var_UCase(VARIANT* String)
{
VARIANT vResult;
	VariantChangeType(&vResult,String,0,VT_BSTR);
	wcsupr(V_BSTR(&vResult));
	return vResult;
}

VARIANT WINAPI VBA_StrComp(VARIANT* String1,VARIANT* String2,VBA_VbCompareMethod Compare)
{
VARIANT vString1,vString2,vResult;
int cmp;

	VariantChangeType(&vString1,String1,0,VT_BSTR);
	VariantChangeType(&vString2,String2,0,VT_BSTR);
	if (Compare == -1)
		Compare = VBA_VbCompareMethod_vbBinaryCompare; /* s/b Option Compare value!!! */
	switch (Compare)
	{
	case VBA_VbCompareMethod_vbBinaryCompare:
		cmp = wcscmp(V_BSTR(&vString1),V_BSTR(&vString2));
		break;
	case VBA_VbCompareMethod_vbTextCompare:
		cmp = wcsicmp(V_BSTR(&vString1),V_BSTR(&vString2));
		break;
	}
	VariantClear(&vString1);
	VariantClear(&vString2);
	switch (Compare)
	{
	case VBA_VbCompareMethod_vbDatabaseCompare:
		diag(ERR_NOT_IMPLEMENTED);
	default:
		diag(ERR_INVALID_FUNCTION_VALUE);
	}
	V_VT(&vResult) = VT_I2;
	if (cmp > 0)
		V_I2(&vResult) = 1;
	else if (cmp < 0)
		V_I2(&vResult) = -1;
	else
		V_I2(&vResult) = 0;
	return vResult;
}

BSTR WINAPI VBA__B_str_Format(VARIANT* Expression,VARIANT* Format,VBA_VbDayOfWeek FirstDayOfWeek,VBA_VbFirstWeekOfYear FirstWeekOfYear)
{
VARIANT vFormat;
BSTR bstr;
	VariantChangeType(&vFormat,Format,0,VT_BSTR);
	VarFormat(Expression,V_BSTR(&vFormat),FirstDayOfWeek,FirstWeekOfYear,0,&bstr);
	VariantClear(&vFormat);
	return bstr;
}

VARIANT WINAPI VBA__B_var_Format(VARIANT* Expression,VARIANT* Format,VBA_VbDayOfWeek FirstDayOfWeek,VBA_VbFirstWeekOfYear FirstWeekOfYear)
{
VARIANT vFormat,vResult;
	VariantChangeType(&vFormat,Format,0,VT_BSTR);
	V_VT(&vResult) = VT_BSTR;
	VarFormat(Expression,V_BSTR(&vFormat),FirstDayOfWeek,FirstWeekOfYear,0,&V_BSTR(&vResult));
	VariantClear(&vFormat);
	return vResult;
}

VARIANT WINAPI VBA_Len(VARIANT* Expression)
{
VARIANT vExpression,vResult;
	VariantChangeType(&vExpression,Expression,0,VT_BSTR);
	V_VT(&vResult) = VT_I4;
	V_I4(&vResult) = SysStringLen(V_BSTR(&vExpression));
	VariantClear(&vExpression);
	return vResult;
}

VARIANT WINAPI VBA_LenB(VARIANT* Expression)
{
VARIANT vExpression,vResult;
	VariantChangeType(&vExpression,Expression,0,VT_BSTR);
	V_VT(&vResult) = VT_I4;
	V_I4(&vResult) = SysStringByteLen(V_BSTR(&vExpression));
	VariantClear(&vExpression);
	return vResult;
}

unsigned char WINAPI VBA_AscB(BSTR String)
{
	return ((CHAR *)String)[0]; /* use WideCharToMultiByte(CP_ACP,0,String,1,buf,sizeof(buf),NULL,NULL); ? */
}

BSTR WINAPI VBA__B_str_ChrB(unsigned char CharCode)
{
BSTR bstr;
	bstr = SysAllocStringLen(NULL,1);
	((CHAR *)bstr)[0] = CharCode;
	((CHAR *)bstr)[1] = 0;
	bstr[1] = 0;
	return bstr;
}

VARIANT WINAPI VBA__B_var_ChrB(unsigned char CharCode)
{
VARIANT vResult;
	V_VT(&vResult) = VT_BSTR;
	V_BSTR(&vResult) = SysAllocStringLen(NULL,1);
	((CHAR *)V_BSTR(&vResult))[0] = CharCode;
	((CHAR *)V_BSTR(&vResult))[1] = 0;
	V_BSTR(&vResult)[1] = 0;
	return vResult;
}

short WINAPI VBA_AscW(BSTR String)
{
	return String[0];
}

BSTR WINAPI VBA__B_str_ChrW(long CharCode)
{
BSTR bstr;
	bstr = SysAllocStringLen(NULL,1);
	bstr[0] = (OLECHAR)CharCode;
	bstr[1] = 0;
	return bstr;
}

VARIANT WINAPI VBA__B_var_ChrW(long CharCode)
{
VARIANT vResult;
	V_VT(&vResult) = VT_BSTR;
	V_BSTR(&vResult) = SysAllocStringLen(NULL,1);
	V_BSTR(&vResult)[0] = (OLECHAR)CharCode;
	V_BSTR(&vResult)[1] = 0;
	return vResult;
}

BSTR WINAPI VBA_FormatDateTime(VARIANT* Expression,VBA_VbDateTimeFormat NamedFormat)
{
BSTR bstr;
	VarFormatDateTime(Expression,NamedFormat,0,&bstr);
	return bstr;
}

BSTR WINAPI VBA_FormatNumber(VARIANT* Expression,INT NumDigitsAfterDecimal,VBA_VbTriState IncludeLeadingDigit,VBA_VbTriState UseParensForNegativeNumbers,VBA_VbTriState GroupDigits)
{
BSTR bstr;
	VarFormatNumber(Expression,NumDigitsAfterDecimal,IncludeLeadingDigit,UseParensForNegativeNumbers,GroupDigits,0,&bstr);
	return bstr;
}

BSTR WINAPI VBA_FormatPercent(VARIANT* Expression,INT NumDigitsAfterDecimal,VBA_VbTriState IncludeLeadingDigit,VBA_VbTriState UseParensForNegativeNumbers,VBA_VbTriState GroupDigits)
{
BSTR bstr;
	VarFormatPercent(Expression,NumDigitsAfterDecimal,IncludeLeadingDigit,UseParensForNegativeNumbers,GroupDigits,0,&bstr);
	return bstr;
}

BSTR WINAPI VBA_FormatCurrency(VARIANT* Expression,INT NumDigitsAfterDecimal,VBA_VbTriState IncludeLeadingDigit,VBA_VbTriState UseParensForNegativeNumbers,VBA_VbTriState GroupDigits)
{
BSTR bstr;
	VarFormatCurrency(Expression,NumDigitsAfterDecimal,IncludeLeadingDigit,UseParensForNegativeNumbers,GroupDigits,0,&bstr);
	return bstr;
}

BSTR WINAPI VBA_WeekdayName(INT Weekday,short Abbreviate,VBA_VbDayOfWeek FirstDayOfWeek)
{
BSTR bstr;
	VarWeekdayName(Weekday,Abbreviate,FirstDayOfWeek,0,&bstr);
	return bstr;
}

BSTR WINAPI VBA_MonthName(INT Month,short Abbreviate)
{
BSTR bstr;
	VarMonthName(Month,Abbreviate,0,&bstr);
	return bstr;
}

BSTR WINAPI VBA_Replace(BSTR Expression,BSTR Find,BSTR Replace,long Start,long Count,VBA_VbCompareMethod Compare)
{
	diag(ERR_NOT_IMPLEMENTED);
}

BSTR WINAPI VBA_StrReverse(BSTR Expression)
{
	diag(ERR_NOT_IMPLEMENTED);
}

BSTR WINAPI VBA_Join(VARIANT* SourceArray,VARIANT* Delimiter)
{
	diag(ERR_NOT_IMPLEMENTED);
}

VARIANT WINAPI VBA_Filter(VARIANT* SourceArray,BSTR Match,short Include,VBA_VbCompareMethod Compare)
{
	diag(ERR_NOT_IMPLEMENTED);
}

long WINAPI VBA_InStrRev(BSTR StringCheck,BSTR StringMatch,long Start,VBA_VbCompareMethod Compare)
{
	diag(ERR_NOT_IMPLEMENTED);
}

VARIANT WINAPI VBA_Split(BSTR Expression,VARIANT* Delimiter,long Limit,VBA_VbCompareMethod Compare)
{
	diag(ERR_NOT_IMPLEMENTED);
}

/* Declaration: Conversion  0 . 0  */
BSTR WINAPI VBA__B_str_Hex(VARIANT* Number)
{
OLECHAR wbuf[32];
VARIANT vNumber;
	VariantChangeType(&vNumber,Number,0,VT_I4); /* is VT_I4 best choice? */
	swprintf(wbuf,L"%lx",V_I4(&vNumber));
/*	VariantClear(&vNumber);*/ /* not needed for VT_I4 */
	return SysAllocString(wbuf);
}

VARIANT WINAPI VBA__B_var_Hex(VARIANT* Number)
{
OLECHAR wbuf[32];
VARIANT vNumber,vResult;
	VariantChangeType(&vNumber,Number,0,VT_I4); /* is VT_I4 best choice? */
	swprintf(wbuf,L"%lx",V_I4(&vNumber));
/*	VariantClear(&vNumber);*/ /* not needed for VT_I4 */
	V_VT(&vResult) = VT_BSTR;
	V_BSTR(&vResult) = SysAllocString(wbuf);
	return vResult;
}

BSTR WINAPI VBA__B_str_Oct(VARIANT* Number)
{
OLECHAR wbuf[32];
VARIANT vNumber;
	VariantChangeType(&vNumber,Number,0,VT_I4); /* is VT_I4 best choice? */
	swprintf(wbuf,L"%lo",V_I4(&vNumber));
/*	VariantClear(&vNumber);*/ /* not needed for VT_I4 */
	return SysAllocString(wbuf);
}

VARIANT WINAPI VBA__B_var_Oct(VARIANT* Number)
{
OLECHAR wbuf[32];
VARIANT vNumber,vResult;
	VariantChangeType(&vNumber,Number,0,VT_I4); /* is VT_I4 best choice? */
	swprintf(wbuf,L"%lo",V_I4(&vNumber));
/*	VariantClear(&vNumber);*/ /* not needed for VT_I4 */
	V_VT(&vResult) = VT_BSTR;
	V_BSTR(&vResult) = SysAllocString(wbuf);
	return vResult;
}

long WINAPI VBA_MacID(BSTR Constant) /* Convert Macintosh 4 character file id to long */
{
	return ((long)Constant[0] << 24) | ((long)Constant[1] << 16) | ((long)Constant[2] << 8) | (long)Constant[3];
}

BSTR WINAPI VBA__B_str_Str(VARIANT* Number)
{
WCHAR wbuf[512];
INT len;
VARTYPE vt;
BSTR bstr;
	vt = V_VT(Number);
	switch(vt & VT_TYPEMASK)
	{
	case VT_EMPTY:
		*wbuf = 0;
		len = 0;
		break;
	case VT_I2:
		len = swprintf(wbuf,L"% d",vt & VT_BYREF ? *V_I2REF(Number) : V_I2(Number));
		break;
	case VT_I4:
		len = swprintf(wbuf,L"% d",vt & VT_BYREF ? *V_I4REF(Number) : V_I4(Number));
		break;
	case VT_R4:
		len = swprintf(wbuf,L"% .*lG",VB_SIGNIFICANCE,vt & VT_BYREF ? (DOUBLE)*V_R4REF(Number) : (DOUBLE)V_R4(Number));
		break;
	case VT_R8:
		len = swprintf(wbuf,L"% .*lG",VB_SIGNIFICANCE,vt & VT_BYREF ? *V_R8REF(Number) : V_R8(Number));
		break;
	case VT_BOOL:
		bstr = SysAllocString((vt & VT_BYREF ? *V_BOOLREF(Number) : V_BOOL(Number)) ? L"True" : L"False");
		goto ret;
		break;
	case VT_UI1:
		len = swprintf(wbuf,L" %u",vt & VT_BYREF ? *V_UI1REF(Number) : V_UI1(Number));
		break;
	case VT_DATE:
		VarBstrFromDate(vt & VT_BYREF ? *V_DATEREF(Number) : V_DATE(Number),0,0,&bstr);
		goto ret;
		break;
	case VT_CY:
		{
		CY cy;
		cy = vt & VT_BYREF ? *V_CYREF(Number) : V_CY(Number);
#if __MSC__ || __WC__
		len = swprintf(wbuf,L"% .05I64d",cy.int64);
#else
		len = swprintf(wbuf,L"% .05lld",cy.int64);
#endif
		wmemmove(wbuf+len-3,wbuf+len-4,4);
		wbuf[len-4] = L'.';
		while (wbuf[len--] == L'0')
			;
		if (wbuf[len] == L'.')
			--len;
		wbuf[++len] = 0;
		}
		break;
	case VT_BSTR:
		bstr = SysAllocString(V_BSTR(Number));
		goto ret;
		break;
	default:
		diag(ERR_NOT_IMPLEMENTED);
	}
	bstr = SysAllocString(wbuf);
ret:
	return bstr;
}

VARIANT WINAPI VBA__B_var_Str(VARIANT* Number)
{
VARIANT vResult;
	V_VT(&vResult) = VT_BSTR;
	V_BSTR(&vResult) = VBA__B_str_Str(Number);
	return vResult;
}

double WINAPI VBA_Val(BSTR String)
{
double d;
	VarR8FromStr(String,0,0,&d);
	return d;
}

BSTR WINAPI VBA_CStr(VARIANT* Expression)
{
VARIANT v;
	VariantInit(&v);
	VariantChangeType(&v,Expression,0,VT_BSTR);
	return V_BSTR(&v);
}

unsigned char WINAPI VBA_CByte(VARIANT* Expression)
{
VARIANT v;
	VariantInit(&v);
	VariantChangeType(&v,Expression,0,VT_UI1);
	return V_UI1(&v);
}

short WINAPI VBA_CBool(VARIANT* Expression)
{
VARIANT v;
	VariantInit(&v);
	VariantChangeType(&v,Expression,0,VT_BOOL);
	return V_BOOL(&v);
}

CY WINAPI VBA_CCur(VARIANT* Expression)
{
VARIANT v;
	VariantInit(&v);
	VariantChangeType(&v,Expression,0,VT_CY);
	return V_CY(&v);
}

DATE WINAPI VBA_CDate(VARIANT* Expression)
{
VARIANT v;
	VariantInit(&v);
	VariantChangeType(&v,Expression,0,VT_DATE);
	return V_DATE(&v);
}

VARIANT WINAPI VBA_CVDate(VARIANT* Expression)
{
VARIANT v;
	VariantInit(&v);
	VariantChangeType(&v,Expression,0,VT_DATE);
	return v;
}

short WINAPI VBA_CInt(VARIANT* Expression)
{
VARIANT v;
	VariantInit(&v);
	VariantChangeType(&v,Expression,0,VT_I2);
	return V_I2(&v);
}

long WINAPI VBA_CLng(VARIANT* Expression)
{
VARIANT v;
	VariantInit(&v);
	VariantChangeType(&v,Expression,0,VT_I4);
	return V_I4(&v);
}

float WINAPI VBA_CSng(VARIANT* Expression)
{
VARIANT v;
	VariantInit(&v);
	VariantChangeType(&v,Expression,0,VT_R4);
	return V_R4(&v);
}

double WINAPI VBA_CDbl(VARIANT* Expression)
{
VARIANT Dest;
	VariantInit(&Dest);
	if (VariantChangeType(&Dest,Expression,0,VT_R8))
		abort();
	return(V_R8(&Dest));
}

VARIANT WINAPI VBA_CVar(VARIANT* Expression) /* needs checking */
{
VARIANT v;
	VariantInit(&v);
	if (V_VT(Expression) != VT_BSTR) /* actually want numerics */
		VariantChangeType(&v,Expression,0,VT_R8);
	return v;
}

VARIANT WINAPI VBA_CVErr(VARIANT* Expression)
{
VARIANT v;
	VariantInit(&v);
	VariantChangeType(&v,Expression,0,VT_ERROR);
	return v;
}

BSTR WINAPI VBA__B_str_Error(VARIANT* ErrorNumber)
{
VARIANT v;
OLECHAR wbuf[512];
	VariantInit(&v);
	VariantChangeType(&v,ErrorNumber,0,VT_ERROR);
	FormatMessage(
		FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS | FORMAT_MESSAGE_MAX_WIDTH_MASK,
		NULL,
		V_ERROR(&v),
		MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), /* Default language */
		wbuf,
		sizeof(wbuf),
		NULL);
/*	VariantClear(&v);*/ /* not needed for VT_ERROR */
	return SysAllocString(wbuf);
}

VARIANT WINAPI VBA__B_var_Error(VARIANT* ErrorNumber)
{
VARIANT vResult;
	V_VT(&vResult) = VT_BSTR;
	V_BSTR(&vResult) = VBA__B_str_Error(ErrorNumber);
	return vResult;
}

VARIANT WINAPI VBA_Fix(VARIANT* Number)
{
VARIANT vResult;
	VariantInit(&vResult);
	if (VarFix(Number,&vResult))
		abort();
	return vResult;
}

VARIANT WINAPI VBA_Int(VARIANT* Number)
{
VARIANT vResult;
	VariantInit(&vResult);
	if (VarInt(Number,&vResult))
		abort();
	return vResult;
}

VARIANT WINAPI VBA_CDec(VARIANT* Expression)
{
VARIANT Dest;
	VariantInit(&Dest);
	if (VariantChangeType(&Dest,Expression,0,VT_DECIMAL))
		abort();
	return Dest;
}

/* Declaration: FileSystem  0 . 0  */
void WINAPI VBA_ChDir(BSTR Path)
{
	_wchdir(Path);
}

void WINAPI VBA_ChDrive(BSTR Drive)
{
	_chdrive(toupper(Drive[0])-'@'); /* makes some assumptions about character set */
}

short WINAPI VBA_EOF(short FileNumber)
{
	return ioeof(FileNumber) ? VARIANT_TRUE : VARIANT_FALSE;
}

long WINAPI VBA_FileAttr(short FileNumber,short ReturnType)
{
	return iofileattr(FileNumber,ReturnType);
}

void WINAPI VBA_FileCopy(BSTR Source,BSTR Destination)
{
CHAR frombuf[MAX_PATH], tobuf[MAX_PATH];

	frombuf[WideCharToMultiByte(CP_ACP,0,Source,-1,frombuf,sizeof(frombuf)-1,NULL,NULL)] = 0;
	tobuf[WideCharToMultiByte(CP_ACP,0,Destination,-1,tobuf,sizeof(tobuf)-1,NULL,NULL)] = 0;
	iocopy(frombuf,strlen(frombuf),tobuf,strlen(tobuf));
}

VARIANT WINAPI VBA_FileDateTime(BSTR PathName)
{
VARIANT vResult;
struct _stat statbuf;
	if (_wstat(PathName,&statbuf))
		diag(ERR_CANNOT_STAT_FILE);
	V_VT(&vResult) = VT_DATE;
	V_DATE(&vResult) = tmtovbdate(localtime(&statbuf.st_mtime));
	return vResult;
}

long WINAPI VBA_FileLen(BSTR PathName)
{
struct _stat statbuf;
	if (_wstat(PathName,&statbuf))
		diag(ERR_CANNOT_STAT_FILE);
	return statbuf.st_size;
}

VBA_VbFileAttribute WINAPI VBA_GetAttr(BSTR PathName)
{
DWORD dwFileAttributes;
	if ((dwFileAttributes = GetFileAttributes(PathName)) == -1)
		diag(ERR_CANNOT_STAT_FILE); /* s/b cannot get file attributes */
	return dwFileAttributes;
}

void WINAPI VBA_Kill(VARIANT* PathName)
{
VARIANT v;
	VariantChangeType(&v,PathName,0,VT_BSTR);
	_wremove(V_BSTR(&v));
	VariantClear(&v);
}

long WINAPI VBA_Loc(short FileNumber)
{
long l;
	fb_tell(iofclu(FileNumber),&l); /* input, output pos needs to be divided by 128 */
	return l;
}

long WINAPI VBA_LOF(short FileNumber)
{
struct stat statbuf;
	fb_stat(iofclu(FileNumber),&statbuf);
	return statbuf.st_size;
}

void WINAPI VBA_MkDir(BSTR Path)
{
#if WIN32
	if (_wmkdir(Path) && errno != EEXIST)
#else
	if (_wmkdir(Path,0700) && errno != EEXIST)
#endif
		diag(ERR_CANNOT_CREATE_FILE); /* s/b cannot make directory */
}

void WINAPI VBA_Reset()
{
	ioclsall();
}

void WINAPI VBA_RmDir(BSTR Path)
{
	if (_wrmdir(Path))
		diag(ERR_CANNOT_REMOVE_FILE); /* s/b cannot remove directory */
}

long WINAPI VBA_Seek(short FileNumber)
{
long l;
	fb_tell(iofclu(FileNumber),&l);
	return l;
}

void WINAPI VBA_SetAttr(BSTR PathName,VBA_VbFileAttribute Attributes)
{
	if (!SetFileAttributes(PathName,Attributes))
		diag(ERR_CANNOT_STAT_FILE); /* s/b cannot get file attributes */
}

BSTR WINAPI VBA__B_str_CurDir(VARIANT* Drive)
{
WCHAR wbuf[MAX_PATH];
int drive;
	if (V_VT(Drive) == VT_BSTR && V_BSTR(Drive) != NULL)
		{
		drive = V_BSTR(Drive)[0] ? toupper(V_BSTR(Drive)[0])-'@' : 0; /* makes assumptions about characterset */
		}
	else
		drive = 0;
	if (_wgetdcwd(drive,wbuf,sizeof(wbuf)) == NULL)
		diag(ERR_CANNOT_OPEN_FILE); /* s/b cannot change directory */
	return SysAllocString(wbuf);
}

VARIANT WINAPI VBA__B_var_CurDir(VARIANT* Drive)
{
VARIANT vDrive;
	V_VT(&vDrive) = VT_BSTR;
	V_BSTR(&vDrive) = VBA__B_str_CurDir(Drive);
	return vDrive;
}

short WINAPI VBA_FreeFile(VARIANT* RangeNumber)
{
#ifdef NEVER
VARIANT v;
FILBLK *fb;
FC fc;
	VariantChangeType(&v,RangeNumber,0,VT_I2);
	fc = V_I2(&v)*256;
/*	VariantClear(&v);*/ /* not needed for VT_I2 */
	do
		for(fc++,fb=bg.x_iob;fb != bg.x_ioe && fc != fb->fb_fc;fb++)
			;
	while(fb != bg.x_ioe && fc < 512);
	if (fc == 512)
		diag(ERR_OUT_OF_FILE_SPACE);
	return fc;
#else
	diag(ERR_NOT_IMPLEMENTED);
#endif
}

BSTR WINAPI VBA_Dir(VARIANT* PathName,VBA_VbFileAttribute Attributes)
{
BSTR bstr;
VARIANT vPathName;
WCHAR *w;
struct _wfinddata_t fileinfo;
static LONG handle = -1;
	
	VariantChangeType(&vPathName,PathName,0,VT_BSTR);
	vbt_printf("VBA_Dir: PathName: vt=%x bstr=%lx\n",V_VT(PathName),V_BSTR(&vPathName));
	do
		if (V_VT(&vPathName) == VT_BSTR && (bstr = V_BSTR(&vPathName)) != NULL && SysStringLen(bstr) != 0)
		{
			if (handle != -1)
				_findclose(handle); /* return isn't relevant */
			vbt_printf("rtcdir: bstr=\"%S\"\n",bstr);
#if WIN32 /* do not include unix code - this is wrong symbol !!! */
			if ((handle = _wfindfirst(*bstr ? bstr : L"*",&fileinfo)) == -1)
#else
			if ((handle = _wfindfirst(*bstr ? bstr : L".",&fileinfo)) == -1)
#endif
				diag(ERR_CANNOT_OPEN_FILE); /* s/b cannot open diretory */
			vbt_printf("rtcdir: name=\"%S\"\n",fileinfo.name);
			w = fileinfo.name;
		}
		else
		{
			if (handle == -1)
				diag(ERR_INVALID_FUNCTION_VALUE);
			if (_wfindnext(handle,&fileinfo))
			{
				w = L"";
				_findclose(handle); /* return isn't relevant */
				handle = -1;
				break;
			}
			vbt_printf("rtcdir: name=\"%S\"\n",fileinfo.name);
			w = fileinfo.name;
		}
	while(Attributes && !(Attributes & fileinfo.attrib));
	return SysAllocString(w);
}

/* Declaration: DateTime  0 . 0  */
VARIANT WINAPI VBA__B_var_DateGet()
{
SYSTEMTIME st;
VARIANT vResult;
	GetSystemTime(&st);
	V_VT(&vResult) = VT_DATE;
	SystemTimeToVariantTime(&st,&V_DATE(&vResult));
	return vResult;
}

void WINAPI VBA__B_str_DateLet(BSTR putval)
{
DATE date;
SYSTEMTIME st;
	VarDateFromStr(putval,0,0,&date);
	VariantTimeToSystemTime(date,&st);
	SetSystemTime(&st);
}

void WINAPI VBA__B_var_DateLet(VARIANT putval)
{
VARIANT v;
SYSTEMTIME st;
	VariantInit(&v);
	VariantChangeType(&v,&putval,0,VT_DATE);
	VariantTimeToSystemTime(V_DATE(&v),&st);
	SetSystemTime(&st);
}

BSTR WINAPI VBA__B_str_DateGet()
{
SYSTEMTIME st;
DATE date;
BSTR bstr;
	GetSystemTime(&st);
	SystemTimeToVariantTime(&st,&date);
	VarBstrFromDate(date,0,0,&bstr);
	return bstr;
}

VARIANT WINAPI VBA_DateSerial(short Year,short Month,short Day)
{
time_t t;
struct tm *tm;
VARIANT vResult;

	t = 0;
	tm = gmtime(&t);
	tm->tm_year = Year-1900;
	tm->tm_mon = Month-1;
	tm->tm_mday = Day;
	V_VT(&vResult) = VT_DATE;
	V_DATE(&vResult) = tmtovbdate(tm);
	return vResult;
}

VARIANT WINAPI VBA_DateValue(BSTR Date)
{
VARIANT vResult;
	V_VT(&vResult) = VT_DATE;
	VarDateFromStr(Date,0,0,&V_DATE(&vResult));
	return vResult;
}

VARIANT WINAPI VBA_Day(VARIANT* Date)
{
VARIANT vDate,vResult;
struct tm *tm;

	if (V_VT(Date) == VT_NULL)
		V_VT(&vResult) = VT_NULL;
	else
		{
		VariantInit(&vDate);
		VariantChangeType(&vDate,Date,0,VT_DATE);
		tm = vbdatetotm(V_DATE(&vDate));
		V_VT(&vResult) = VT_I2;
		V_I2(&vResult) = (INT16)tm->tm_mday;
		}
	return vResult;
}

VARIANT WINAPI VBA_Hour(VARIANT* Time)
{
VARIANT vTime,vResult;
struct tm *tm;

	if (V_VT(Time) == VT_NULL)
		V_VT(&vResult) = VT_NULL;
	else
		{
		VariantInit(&vTime);
		VariantChangeType(&vTime,Time,0,VT_DATE);
		tm = vbdatetotm(V_DATE(&vTime));
		V_VT(&vResult) = VT_I2;
		V_I2(&vResult) = (INT16)tm->tm_hour;
		}
	return vResult;
}

VARIANT WINAPI VBA_Minute(VARIANT* Time)
{
VARIANT vTime,vResult;
struct tm *tm;

	if (V_VT(Time) == VT_NULL)
		V_VT(&vResult) = VT_NULL;
	else
		{
		VariantInit(&vTime);
		VariantChangeType(&vTime,Time,0,VT_DATE);
		tm = vbdatetotm(V_DATE(&vTime));
		V_VT(&vResult) = VT_I2;
		V_I2(&vResult) = (INT16)tm->tm_min;
		}
	return vResult;
}

VARIANT WINAPI VBA_Month(VARIANT* Date)
{
VARIANT vDate,vResult;
struct tm *tm;

	if (V_VT(Date) == VT_NULL)
		V_VT(&vResult) = VT_NULL;
	else
		{
		VariantInit(&vDate);
		VariantChangeType(&vDate,Date,0,VT_DATE);
		tm = vbdatetotm(V_DATE(&vDate));
		V_VT(&vResult) = VT_I2;
		V_I2(&vResult) = (INT16)tm->tm_mon+1;
		}
	return vResult;
}

VARIANT WINAPI VBA_NowGet()
{
SYSTEMTIME st;
VARIANT vResult;
	GetLocalTime(&st);
	V_VT(&vResult) = VT_DATE;
	SystemTimeToVariantTime(&st,&V_DATE(&vResult));
	return vResult;
}

VARIANT WINAPI VBA_Second(VARIANT* Time)
{
VARIANT vTime,vResult;
struct tm *tm;

	if (V_VT(Time) == VT_NULL)
		V_VT(&vResult) = VT_NULL;
	else
		{
		VariantInit(&vTime);
		VariantChangeType(&vTime,Time,0,VT_DATE);
		tm = vbdatetotm(V_DATE(&vTime));
		V_VT(&vResult) = VT_I2;
		V_I2(&vResult) = (INT16)tm->tm_sec;
		}
	return vResult;
}

VARIANT WINAPI VBA__B_var_TimeGet()
{
SYSTEMTIME st;
VARIANT vResult;
	GetSystemTime(&st);
	V_VT(&vResult) = VT_DATE;
	SystemTimeToVariantTime(&st,&V_DATE(&vResult));
	return vResult;
}

void WINAPI VBA__B_str_TimeLet(BSTR putval)
{
DATE date;
SYSTEMTIME st;
	VarDateFromStr(putval,0,0,&date);
	VariantTimeToSystemTime(date,&st);
	SetSystemTime(&st);
}

void WINAPI VBA__B_var_TimeLet(VARIANT putval)
{
VARIANT v;
SYSTEMTIME st;
	VariantInit(&v);
	VariantChangeType(&v,&putval,0,VT_DATE);
	VariantTimeToSystemTime(V_DATE(&v),&st);
	SetSystemTime(&st);
}

BSTR WINAPI VBA__B_str_TimeGet()
{
SYSTEMTIME st;
DATE date;
BSTR bstr;
	GetSystemTime(&st);
	SystemTimeToVariantTime(&st,&date);
	VarBstrFromDate(date,0,0,&bstr);
	return bstr;
}

float WINAPI VBA_TimerGet()
{
	return((float)((double)clock()/(double)CLOCKS_PER_SEC));
}

VARIANT WINAPI VBA_TimeSerial(short Hour,short Minute,short Second)
{
VARIANT vResult;
time_t t;
struct tm *tm;

	t = 0;
	tm = gmtime(&t);
	tm->tm_hour = Hour;
	tm->tm_min = Minute;
	tm->tm_sec = Second;
	t = (time_t)tm->tm_hour*3600+(time_t)tm->tm_min*60+(time_t)tm->tm_sec;
	if (mktime(tm) == -1)
		diag(ERR_NOT_IMPLEMENTED);
	V_VT(&vResult) = VT_DATE;
	V_DATE(&vResult) = (DATE)t/(DATE)SW_SECS_PER_DAY;
	return vResult;
}

VARIANT WINAPI VBA_TimeValue(BSTR Time)
{
VARIANT vResult;
DATE date;
	VarDateFromStr(Time,0,0,&date);
	V_VT(&vResult) = VT_DATE;
	V_DATE(&vResult) = date-(DATE)(long)date; /* remove fractional part */
	return vResult;
}

VARIANT WINAPI VBA_Weekday(VARIANT* Date,VBA_VbDayOfWeek FirstDayOfWeek)
{
VARIANT vDate,vResult;
struct tm *tm;

	if (FirstDayOfWeek < 0 || FirstDayOfWeek > 7)
		diag(ERR_INVALID_FUNCTION_VALUE);
	if (FirstDayOfWeek == 0)
		FirstDayOfWeek = 1; /* if vbUseSystem then use Sunday as default */ /* need to use locale default */
	if (V_VT(Date) == VT_NULL)
		V_VT(&vResult) = VT_NULL;
	else
		{
		VariantInit(&vDate);
		VariantChangeType(&vDate,Date,0,VT_DATE);
		tm = vbdatetotm(V_DATE(&vDate));
		V_VT(&vResult) = VT_I2;
		V_I2(&vResult) = (INT16)((tm->tm_wday+(8-FirstDayOfWeek)) % 7 + 1);
		}
	return vResult;
}

VARIANT WINAPI VBA_Year(VARIANT* Date)
{
VARIANT vDate,vResult;
struct tm *tm;

	if (V_VT(Date) == VT_NULL)
		V_VT(&vResult) = VT_NULL;
	else
		{
		VariantInit(&vDate);
		VariantChangeType(&vDate,Date,0,VT_DATE);
		tm = vbdatetotm(V_DATE(&vDate));
		V_VT(&vResult) = VT_I2;
		V_I2(&vResult) = (INT16)tm->tm_year+1900;
		}
	return vResult;
}

VARIANT WINAPI VBA_DateAdd(BSTR Interval,double Number,VARIANT* Date)
{
VARIANT vResult;
struct tm *tm;

	if (V_VT(Date) == VT_NULL)
		V_VT(&vResult) = VT_NULL;
	else
		{
		VariantInit(&vResult);
		VariantChangeType(&vResult,Date,0,VT_DATE);
		if ((tm = vbdatetotm(V_DATE(&vResult))) == NULL)
			diag(ERR_INVALID_FUNCTION_VALUE);
		if (wcscmp(Interval,L"yyyy") == 0)
			{
			tm->tm_year += (int)Number;
			}
		else if (wcscmp(Interval,L"q") == 0)
			{
			tm->tm_mon += (int)Number*3;
			}
		else if (wcscmp(Interval,L"m") == 0)
			{
			tm->tm_mon += (int)Number;
			}
		else if (wcscmp(Interval,L"y") == 0)
			{
			tm->tm_mday += (int)Number; /* "y", "d" and "w" are identical */
			}
		else if (wcscmp(Interval,L"d") == 0)
			{
			tm->tm_mday += (int)Number; /* "y", "d" and "w" are identical */
			}
		else if (wcscmp(Interval,L"w") == 0)
			{
			tm->tm_mday += (int)Number; /* "y", "d" and "w" are identical */
			}
		else if (wcscmp(Interval,L"ww") == 0)
			{
			tm->tm_mday += (int)Number*7;
			}
		else if (wcscmp(Interval,L"h") == 0)
			{
			tm->tm_hour += (int)Number;
			}
		else if (wcscmp(Interval,L"n") == 0)
			{
			tm->tm_min += (int)Number;
			}
		else if (wcscmp(Interval,L"s") == 0)
			{
			tm->tm_sec += (int)Number;
			}
		V_VT(&vResult) = VT_DATE;
		V_DATE(&vResult) = tmtovbdate(tm); /* need to check for errors (negative date) */
		}
	return vResult;
}

VARIANT WINAPI VBA_DateDiff(BSTR Interval,VARIANT* Date1,VARIANT* Date2,VBA_VbDayOfWeek FirstDayOfWeek,VBA_VbFirstWeekOfYear FirstWeekOfYear)
{
VARIANT vResult;
struct tm *tm,tm1,tm2;
int i;
	if (FirstDayOfWeek < 0 || FirstDayOfWeek > 7)
		diag(ERR_INVALID_FUNCTION_VALUE);
	if (FirstDayOfWeek)
		FirstDayOfWeek -= 1; /* if vbUseSystem then use Sunday as default */ /* need to use locale default */
	if (FirstWeekOfYear < 0 || FirstWeekOfYear > 3)
		diag(ERR_INVALID_FUNCTION_VALUE);
	if (V_VT(Date1) == VT_NULL || V_VT(Date2) == VT_NULL)
		V_VT(&vResult) = VT_NULL;
	else
		{
		VariantChangeType(&vResult,Date1,0,VT_DATE);
		if ((tm = vbdatetotm(V_DATE(&vResult))) == NULL)
			diag(ERR_INVALID_FUNCTION_VALUE);
		tm1 = *tm;
		VariantChangeType(&vResult,Date2,0,VT_DATE);
		if ((tm = vbdatetotm(V_DATE(&vResult))) == NULL)
			diag(ERR_INVALID_FUNCTION_VALUE);
		tm2 = *tm;
		if (wcscmp(Interval,L"yyyy") == 0)
			{
			i = tm2.tm_year-tm1.tm_year;
			}
		else if (wcscmp(Interval,L"q") == 0)
			{
			i = tm2.tm_year-tm1.tm_year;
			i *= 4;
			i += tm2.tm_mon/3-tm1.tm_mon/3;
			}
		else if (wcscmp(Interval,L"m") == 0)
			{
			i = tm2.tm_year-tm1.tm_year;
			i *= 12;
			i += tm2.tm_mon-tm1.tm_mon;
			}
		else if (wcscmp(Interval,L"y") == 0)
			{
			i = tm2.tm_year-tm1.tm_year;
			i = (i < 0 ? i-(4-(tm1.tm_year&3)&3):i+(((tm1.tm_year&3)+3)&3))/4+i*365;
			i += tm2.tm_yday-tm1.tm_yday;
			}
		else if (wcscmp(Interval,L"d") == 0)
			{
			i = tm2.tm_year-tm1.tm_year;
			i = (i < 0 ? i-(4-(tm1.tm_year&3)&3):i+(((tm1.tm_year&3)+3)&3))/4+i*365;
			i += tm2.tm_yday-tm1.tm_yday;
			}
		else if (wcscmp(Interval,L"w") == 0) /* contrary to docs, firstdayofweek doesn't seem to be used */
			{
			i = tm2.tm_year-tm1.tm_year;
			i = (i < 0 ? i-(4-(tm1.tm_year&3)&3):i+(((tm1.tm_year&3)+3)&3))/4+i*365;
			i += tm2.tm_yday-tm1.tm_yday;
			i /= 7;
			}
		else if (wcscmp(Interval,L"ww") == 0)
			{  /* contrary to docs, firstweekofyear doesn't seem to be used */
			i = tm2.tm_year-tm1.tm_year;
			i = (i < 0 ? i-(4-(tm1.tm_year&3)&3):i+(((tm1.tm_year&3)+3)&3))/4+i*365;
			i += tm2.tm_yday-tm1.tm_yday;
			if (i > 0)
				i += 6-(7+tm2.tm_wday-FirstDayOfWeek)%7;
			else
				i -= 6-(7+tm1.tm_wday-FirstDayOfWeek)%7;
			i /= 7;
			}
		else if (wcscmp(Interval,L"h") == 0)
			{
			i = tm2.tm_year-tm1.tm_year;
			i = (i < 0 ? i-(4-(tm1.tm_year&3)&3):i+(((tm1.tm_year&3)+3)&3))/4+i*365;
			i += tm2.tm_yday-tm1.tm_yday;
			i *= 24;
			i += tm2.tm_hour-tm1.tm_hour;
			}
		else if (wcscmp(Interval,L"n") == 0)
			{
			i = tm2.tm_year-tm1.tm_year;
			i = (i < 0 ? i-(4-(tm1.tm_year&3)&3):i+(((tm1.tm_year&3)+3)&3))/4+i*365;
			i += tm2.tm_yday-tm1.tm_yday;
			i *= 24;
			i += tm2.tm_hour-tm1.tm_hour;
			i *= 60;
			i += tm2.tm_min-tm1.tm_min;
			}
		else if (wcscmp(Interval,L"s") == 0)
			{
			i = tm2.tm_year-tm1.tm_year;
			i = (i < 0 ? i-(4-(tm1.tm_year&3)&3):i+(((tm1.tm_year&3)+3)&3))/4+i*365;
			i += tm2.tm_yday-tm1.tm_yday;
			i *= 24;
			i += tm2.tm_hour-tm1.tm_hour;
			i *= 60;
			i += tm2.tm_min-tm1.tm_min;
			i *= 60;
			i += tm2.tm_sec-tm1.tm_sec;
			}
		V_VT(&vResult) = VT_I4;
		V_I4(&vResult) = i;
		}
	return vResult;
}

VARIANT WINAPI VBA_DatePart(BSTR Interval,VARIANT* Date,VBA_VbDayOfWeek FirstDayOfWeek,VBA_VbFirstWeekOfYear FirstWeekOfYear)
{
VARIANT vResult;
struct tm *tm;
INT16 i;
	if (FirstDayOfWeek < 0 || FirstDayOfWeek > 7)
		diag(ERR_INVALID_FUNCTION_VALUE);
	if (FirstDayOfWeek == 0) /* use system default */
		FirstDayOfWeek = 1; /* if vbUseSystem then use Sunday as default */ /* need to use locale default */
	if (FirstWeekOfYear < 0 || FirstWeekOfYear > 3)
		diag(ERR_INVALID_FUNCTION_VALUE);
	if (V_VT(Date) == VT_NULL)
		V_VT(&vResult) = VT_NULL;
	else
		{
		VariantChangeType(&vResult,Date,0,VT_DATE);
		if ((tm = vbdatetotm(V_DATE(&vResult))) == NULL)
			diag(ERR_INVALID_FUNCTION_VALUE);
		if (wcscmp(Interval,L"yyyy") == 0)
			{
			i = tm->tm_year+1900;
			}
		else if (wcscmp(Interval,L"q") == 0)
			{
			i = (tm->tm_mon+3)/3;
			}
		else if (wcscmp(Interval,L"m") == 0)
			{
			i = tm->tm_mon+1;
			}
		else if (wcscmp(Interval,L"y") == 0)
			{
			i = tm->tm_yday+1;
			}
		else if (wcscmp(Interval,L"d") == 0)
			{
			i = tm->tm_mday;
			}
		else if (wcscmp(Interval,L"w") == 0)
			{
			i = (tm->tm_wday+(8-FirstDayOfWeek)) % 7 + 1;
			}
		else if (wcscmp(Interval,L"ww") == 0)
			{
			INT jan1wday;
			jan1wday = (tm->tm_wday-tm->tm_yday % 7+15-FirstDayOfWeek) % 7;
			switch(FirstWeekOfYear)
			{
			case 0:
				/* if vbUseSystem then use Jan 1 as default */ /* need to use locale default */
			case 1:
				i = (tm->tm_yday+jan1wday)/7 + 1;
				break;
			case 2:
				if (jan1wday > 3)
					if (tm->tm_yday+jan1wday < 7)
						if (jan1wday == 4 || (jan1wday == 5 && ((tm->tm_year-1) & 3) == 0))
							i = 53;
						else
							i = 52;
					else
						i = (tm->tm_yday+jan1wday)/7;
				else
					i = (tm->tm_yday+jan1wday)/7 + 1;
				break;
			case 3:
				if (jan1wday > 0)
					if (tm->tm_yday+jan1wday < 7)
						if (jan1wday == 1 || (jan1wday == 2 && ((tm->tm_year-1) & 3) == 0))
							i = 53;
						else
							i = 52;
					else
						i = (tm->tm_yday+jan1wday)/7;
				else
					i = (tm->tm_yday+jan1wday)/7 + 1;
				break;
			default:
				diag(ERR_SYSTEM_ERROR);
			}
{
tm->tm_mon = 0;
tm->tm_mday = 1;
mktime(tm);
if (jan1wday != tm->tm_wday)
	diag(ERR_SYSTEM_ERROR);
}
			}
		else if (wcscmp(Interval,L"h") == 0)
			{
			i = tm->tm_hour;
			}
		else if (wcscmp(Interval,L"n") == 0)
			{
			i = tm->tm_min;
			}
		else if (wcscmp(Interval,L"s") == 0)
			{
			i = tm->tm_sec;
			}
		V_VT(&vResult) = VT_I2;
		V_I2(&vResult) = i;
		}
	return vResult;
}

VBA_VbCalendar WINAPI VBA_CalendarGet()
{
	return 0; /* only Gregorian is implemented */
}

void WINAPI VBA_CalendarLet(VBA_VbCalendar putval)
{
	if (putval)
		diag(ERR_NOT_IMPLEMENTED);
}

/* Declaration: Information  0 . 0  */
long WINAPI VBA_Erl() /* hidden - line number of statement causing error */
{
	diag(ERR_NOT_IMPLEMENTED);
}

VBA__ErrObject *WINAPI VBA_Err()
{
#ifdef NEVER
extern IErrObjectVB ErrObjectVB;
	return ErrObjectVB;
#else
	diag(ERR_NOT_IMPLEMENTED);
#endif
}

VBA_VbIMEStatus WINAPI VBA_IMEStatus()
{
	diag(ERR_NOT_IMPLEMENTED);
}

short WINAPI VBA_IsArray(VARIANT* VarName)
{
	return V_VT(VarName) & VT_ARRAY ? VARIANT_TRUE : VARIANT_FALSE;
}

short WINAPI VBA_IsDate(VARIANT* Expression)
{
	return V_VT(Expression) == VT_DATE ? VARIANT_TRUE : VARIANT_FALSE;
}

short WINAPI VBA_IsEmpty(VARIANT* Expression)
{
	return V_VT(Expression) == VT_EMPTY ? VARIANT_TRUE : VARIANT_FALSE;
}

short WINAPI VBA_IsError(VARIANT* Expression)
{
	return V_VT(Expression) == VT_ERROR ? VARIANT_TRUE : VARIANT_FALSE;
}

short WINAPI VBA_IsMissing(VARIANT* ArgName) /* s/b NULL, VT_ERROR or something else? */
{
	return V_VT(ArgName) == VT_NULL ? VARIANT_TRUE : VARIANT_FALSE;
}

short WINAPI VBA_IsNull(VARIANT* Expression)
{
	return V_VT(Expression) == VT_NULL ? VARIANT_TRUE : VARIANT_FALSE;
}

short WINAPI VBA_IsNumeric(VARIANT* Expression)
{
VARIANT v;
/* is there a better way of testing for numeric? */
/* test to be sure that VT_DATE returns false */
	return VariantChangeType(&v,Expression,0,VT_R8) == 0 ? VARIANT_TRUE : VARIANT_FALSE;
}

short WINAPI VBA_IsObject(VARIANT* Expression)
{
	return V_VT(Expression) == VT_DISPATCH ? VARIANT_TRUE : VARIANT_FALSE;
}

/* Isn't there an API to do this? */
/* Need to add more type names to table */
BSTR WINAPI VBA_TypeName(VARIANT* VarName)
{
VARTYPE vt,vtm;
static OLECHAR *vnames[] = { /* This array contains only VB variant types - might want to add more */
/* 00 */        L"Empty",
/* 01 */        L"Null",
/* 02 */        L"Integer",
/* 03 */        L"Long",
/* 04 */        L"Single",
/* 05 */        L"Double",
/* 06 */        L"Currency",
/* 07 */        L"Date",
/* 08 */        L"String",
/* 09 */        L"Object",
/* 0A */        L"Error",
/* 0B */        L"Boolean",
/* 0C */        L"Variant",
/* 0D */        NULL,
/* 0E */        L"Decimal",
/* 0F */        NULL,
/* 10 */        NULL,
/* 11 */        L"Byte"
};
	vt = V_VT(VarName);
	vtm = vt & VT_TYPEMASK;
	if (vtm > sizeof(vnames)/sizeof(vnames[0]))
		diag(ERR_NOT_IMPLEMENTED);
	if (vnames[vtm] == NULL)
		diag(ERR_NOT_IMPLEMENTED);
	if (vt & VT_ARRAY)
		return wcscat(SysAllocStringLen(vnames[vtm],SysStringLen(vnames[vtm]+2)),L"()");
	else
		return SysAllocString(vnames[vtm]);
}

VBA_VbVarType WINAPI VBA_VarType(VARIANT* VarName)
{
	return V_VT(VarName);
}

long WINAPI VBA_QBColor(short Color)
{
	diag(ERR_NOT_IMPLEMENTED);
}

long WINAPI VBA_RGB(short Red,short Green,short Blue)
{
	diag(ERR_NOT_IMPLEMENTED);
}

/* Declaration: Interaction  0 . 0  */
void WINAPI VBA_AppActivate(VARIANT* Title,VARIANT* Wait)
{
#ifdef NEVER
	EnumWindows(EnumFunc,Title);
	SetForegroundWindow(hWnd); /* put in EnumFunc? */
#else
	diag(ERR_NOT_IMPLEMENTED);
#endif
}

void WINAPI VBA_Beep()
{
	MessageBeep(MB_OK);
}

VARIANT WINAPI VBA_CreateObject(BSTR Class, BSTR ServerName)
{
VARIANT v;
HRESULT hr;
CLSID clsid;
WCHAR *w;
COSERVERINFO ServerInfo;
MULTI_QI mqi[1];

	vbt_printf("\tVBA_CreateObject: Class=%S ServerName=%S\n",Class,ServerName);
	VariantInit(&v);
	V_VT(&v) = VT_ERROR; /* ??? */

	if (FAILED(hr = CLSIDFromString(Class,&clsid)))
		{
		vbt_printf("\tCLSIDFromString: hr=%x\n",hr);
        	hr = CLSIDFromProgID(Class,&clsid);
		vbt_printf("\tCLSIDFromProgID: hr=%x\n",hr);
		if (FAILED(hr))
			return(v);
		}

	if (SUCCEEDED(StringFromCLSID(&clsid,&w)))
		vbt_printf("\tStringFromCLSID=%S\n",w);
        if (SUCCEEDED(ProgIDFromCLSID(&clsid,&w)))
		vbt_printf("\tProgIDFromCLSID=%S\n",w);
	vbt_printf("\tServerName=%S\n",ServerName);

	memset(&ServerInfo,0,sizeof(ServerInfo));
	ServerInfo.pwszName = ServerName;
	mqi[0].pIID = &IID_IUnknown; /*IID_IDispatch;*/
	mqi[0].pItf = NULL;
	mqi[0].hr = S_OK;
/* Note: CoCreateInstance searches registry CLSID key for correct handler */
	hr = CoCreateInstanceEx(&clsid,NULL,CLSCTX_INPROC_SERVER,&ServerInfo,1,mqi);
/* 80004002 means interface (eg, IID_IDispatch) doesn't exist */
        vbt_printf("\tWinCoCreateInstanceEx: INPROC: hr=%x mqi[0].hr=%x ServerName=%.32S *pdisp=%lx\n",hr,mqi[0].hr,ServerName,mqi[0].pItf);
/* CLSCTX_INPROC_SERVER may not handle IUnknown (Excel.Application, hr=80040111)
   (because of threading model), so try LOCAL/REMOTE servers */
	if (FAILED(hr))
	{
		hr = CoCreateInstanceEx(&clsid,NULL,CLSCTX_LOCAL_SERVER | CLSCTX_REMOTE_SERVER,&ServerInfo,1,mqi);
	        vbt_printf("\tWinCoCreateInstanceEx: LOCAL/REMOTE hr=%x mqi[0].hr=%x ServerName=%.32S *pdisp=%lx\n",hr,mqi[0].hr,ServerName,mqi[0].pItf);
	}
	if (SUCCEEDED(hr))
	{
		V_VT(&v) = VT_UNKNOWN;
		V_UNKNOWN(&v) = mqi[0].pItf;
	}
	vbt_printf("\thr=%lx\n",hr);
	return v;
}

short WINAPI VBA_DoEvents()
{
	Sleep(0);	/* use MsgWaitForMultipleObjects instead? */
	return 0;	/* return number of open forms - zero for now */
}

VARIANT WINAPI VBA_GetObject(VARIANT* PathName,VARIANT* Class)
{
#ifdef NEVER
IUnknown *punk;
IDispatch *pdisp;
	if (V_VT(Class) == VT_NULL)
	{
	CLSID clsid;
		CLSIDFromProgID(PathName,&clsid);
		GetActiveObject(&clsid,NULL,&punk);
		IUnknown_QueryInterface(punk,&pdisp);
	}
	else
	{
		...
	}
	return pdisp;
#else
	diag(ERR_NOT_IMPLEMENTED);
#endif
}

BSTR WINAPI VBA_InputBox(VARIANT* Prompt,VARIANT* Title,VARIANT* Default,VARIANT* XPos,VARIANT* YPos,VARIANT* HelpFile,VARIANT* Context)
{
	diag(ERR_NOT_IMPLEMENTED);
}

BSTR WINAPI VBA_MacScript(BSTR Script) /* hidden */
{
	diag(ERR_NOT_IMPLEMENTED);
}

VBA_VbMsgBoxResult WINAPI VBA_MsgBox(VARIANT* Prompt,VBA_VbMsgBoxStyle Buttons,VARIANT* Title,VARIANT* HelpFile,VARIANT* Context)
{
VARIANT bPrompt;
VARIANT bTitle;
	VariantInit(&bPrompt);
	VariantInit(&bTitle);
	VariantChangeType(&bPrompt,Prompt,0,VT_BSTR);
/* fixme: VariantChangeType converts VT_ERROR (missing arg) to VT_EMPTY - want data type default filled */
	if (V_VT(&bPrompt) == VT_EMPTY) V_BSTR(&bPrompt) = L"";
	VariantChangeType(&bTitle,Title,0,VT_BSTR);
	if (V_VT(&bTitle) == VT_EMPTY) V_BSTR(&bTitle) = L"";
	return(MessageBox(0,V_BSTR(&bPrompt),V_BSTR(&bTitle),Buttons));
}

void WINAPI VBA_SendKeys(BSTR String,VARIANT* Wait)
{
#ifdef NEVER
	SendInput(cInputs,pInputs,cbSize);
#else
	diag(ERR_NOT_IMPLEMENTED);
#endif
}

double WINAPI VBA_Shell(VARIANT* PathName,VBA_VbAppWinStyle WindowStyle)
{
#ifdef NEVER
SHELLEXECUTEINFO ExecInfo;
	ExecInfo.cbSize = sizeof(ExecInfo);
	...
	ShellExecuteEx(&ExecInfo);
	return (double)ExecInfo.hProcess; /* but process won't be closed! ok?? */
#else
double d;
	if (V_VT(PathName) != VT_BSTR)
		diag(ERR_INVALID_FUNCTION_VALUE);
	d = (double)(long)ShellExecute(0,NULL,V_BSTR(PathName),NULL,NULL,WindowStyle);
	if (d <= 32.)
		diag(ERR_CANNOT_OPEN_FILE); /* can't shell program */
	return d; /* is this acceptable task id? */
#endif
}

VARIANT WINAPI VBA_Partition(VARIANT* Number,VARIANT* Start,VARIANT* Stop,VARIANT* Interval)
{
	diag(ERR_NOT_IMPLEMENTED);
}

VARIANT WINAPI VBA_Choose(float Index,VARIANT* Choice)
{
long l;
VARIANT vResult;
	if (!(V_VT(Choice) & VT_ARRAY))
		diag(ERR_INVALID_FUNCTION_VALUE); /* expecting array */
	if (V_VT(Choice) != (VT_VARIANT | VT_ARRAY))
		diag(ERR_INVALID_FUNCTION_VALUE); /* expecting array */
	if(SafeArrayGetDim(V_ARRAY(Choice)) != 1)
		diag(ERR_INVALID_FUNCTION_VALUE); /* invalid number of dimensions */
	l = (long)Index; /* why is Index of type float? */
	V_VT(&vResult) = VT_NULL;
	SafeArrayGetElement(V_ARRAY(Choice),&l,&vResult);
	return vResult;
}

VARIANT WINAPI VBA__B_var_Environ(VARIANT* Expression)
{
VARIANT vResult;
	V_VT(&vResult) = VT_BSTR;
	V_BSTR(&vResult) = VBA__B_str_Environ(Expression);
	return vResult;
}

BSTR WINAPI VBA__B_str_Environ(VARIANT* Expression)
{
WCHAR *p;
	p = L"";
	if (V_VT(Expression) == VT_BSTR)
		{
		if (V_BSTR(Expression) != NULL)
			{
			if ((p = _wgetenv(V_BSTR(Expression))) == NULL)
				p = L"";
			}
		}
/* must be some better way of handling type vs. type|BYREF!! */
	else if (V_VT(Expression) == (VT_BSTR | VT_BYREF))
		{
		if (*V_BSTRREF(Expression) != NULL)
			{
			if ((p = _wgetenv(*V_BSTRREF(Expression))) == NULL)
				p = L"";
			}
		}
	else
		{
		VARIANT vExpression;
		INT16 n;
		VariantChangeType(&vExpression,Expression,0,VT_I2);
		if ((n = V_I2(&vExpression)) > 0)
			{
			WCHAR **envp;
			envp = g_wenvp;
			while(--n)
				envp++;
			if (envp != NULL)
				p = *envp;
			}
		}
	return SysAllocString(p);
}

VARIANT WINAPI VBA_Switch(VARIANT* VarExpr)
{
long l,LBound,UBound;
VARIANT vResult;
	if (!(V_VT(VarExpr) & VT_ARRAY))
		diag(ERR_INVALID_FUNCTION_VALUE); /* expecting array */
	if (V_VT(VarExpr) != (VT_VARIANT | VT_ARRAY))
		diag(ERR_INVALID_FUNCTION_VALUE); /* expecting array */
	if(SafeArrayGetDim(V_ARRAY(VarExpr)) != 1)
		diag(ERR_INVALID_FUNCTION_VALUE); /* invalid number of dimensions */
	SafeArrayGetLBound(V_ARRAY(VarExpr),1,&LBound);
	SafeArrayGetUBound(V_ARRAY(VarExpr),1,&UBound);
	for(l=LBound;l<=UBound;l++)
	{
		VARIANT vBool;
		SafeArrayGetElement(V_ARRAY(VarExpr),&l,&vBool);
		l++;
		if (V_BOOL(&vBool))
		{
			SafeArrayGetElement(V_ARRAY(VarExpr),&l,&vResult);
			return vResult;
		}
	}
	V_VT(&vResult) = VT_NULL;
	return vResult;
}

VARIANT WINAPI VBA__B_var_Command()
{
VARIANT vResult;
	V_VT(&vResult) = VT_BSTR;
	V_BSTR(&vResult) = VBA__B_str_Command();
	return vResult;
}

BSTR WINAPI VBA__B_str_Command()
{
WCHAR wbuf[65535]; /* what is max command line size? */
int i;
	for(*wbuf=0,i=2;i<g_argc;++i)
	{
		wcsncat(wbuf,g_wargv[i],sizeof(wbuf));
		if (i+1 < g_argc)
			wcsncat(wbuf,L" ",sizeof(wbuf));
	}
	return SysAllocStringLen(wbuf,wcslen(wbuf));
}

VARIANT WINAPI VBA_IIf(VARIANT* Expression,VARIANT* TruePart,VARIANT* FalsePart)
{
VARIANT vExpression,vResult;
	VariantChangeType(&vExpression,Expression,0,VT_BOOL);
	VariantCopy(&vResult,V_BOOL(&vExpression) ? TruePart : FalsePart);
	return vResult;
}

BSTR WINAPI VBA_GetSetting(BSTR AppName,BSTR Section,BSTR Key,VARIANT Default)
{
BSTR wSubKey;
HKEY hKey;
LONG lResult;
BSTR bstr;

	wSubKey = SysAllocStringLen(AppName,SysStringLen(AppName)+1+SysStringLen(Section));
	if (Section[0])
	{
		wcscat(wSubKey,L"\\");
		wcscat(wSubKey,Section);
	}
	lResult = RegOpenKeyEx(HKEY_CURRENT_USER, wSubKey, 0, KEY_READ, &hKey);
	SysFreeString(wSubKey);
	if(lResult == ERROR_FILE_NOT_FOUND)
	{
	VARIANT vResult;
		lResult = ERROR_SUCCESS;
		VariantInit(&vResult);
		VariantChangeType(&vResult,&Default,0,VT_BSTR);
		bstr = V_BSTR(&vResult);
	}
	else
	{
	BYTE Data[1024]; /* what is max reg entry? */
	DWORD Type;
	DWORD cbData = sizeof(Data); /* yes, in bytes */
		lResult = RegQueryValueEx(hKey,Key,NULL,&Type,Data,&cbData);
		if(lResult == ERROR_SUCCESS)
			switch (Type)
			{
			case REG_DWORD:
				VarBstrFromUI4(*(ULONG *)Data,0,0,&bstr);
				break;
			case REG_SZ:
				bstr = SysAllocString((LPOLESTR)Data); /* REG_SZ is null terminated UNICODE */
				break;
			default:
/* need to convert more REG types to BSTR */
				diag(ERR_INVALID_FUNCTION_VALUE);
			}
	}
	RegCloseKey(hKey);
 	if(lResult != ERROR_SUCCESS)
		diag(ERR_INVALID_FUNCTION_VALUE);
	return bstr;
}

void WINAPI VBA_SaveSetting(BSTR AppName,BSTR Section,BSTR Key,BSTR Setting)
{
BSTR wSubKey;
HKEY hKey;
LONG lResult;

	wSubKey = SysAllocStringLen(AppName,SysStringLen(AppName)+1+SysStringLen(Section));
	if (Section[0])
	{
		wcscat(wSubKey,L"\\");
		wcscat(wSubKey,Section);
	}
	lResult = RegOpenKeyEx(HKEY_CURRENT_USER, wSubKey, 0, KEY_READ, &hKey);
	if(lResult == ERROR_SUCCESS)
		lResult = RegSetValueEx(hKey, Key, 0, REG_SZ, (BYTE*)Setting, SysStringByteLen(Setting)+1);
	RegCloseKey(hKey);
	if(lResult != ERROR_SUCCESS)
		diag(ERR_INVALID_FUNCTION_VALUE);
}

void WINAPI VBA_DeleteSetting(BSTR AppName,VARIANT Section,VARIANT Key)
{
BSTR wSubKey;
HKEY hKey;
LONG lResult;
VARIANT vSection;

	VariantInit(&vSection);
	VariantChangeType(&vSection,&Section,0,VT_BSTR);
	wSubKey = SysAllocStringLen(AppName,SysStringLen(AppName)+1+SysStringLen(V_BSTR(&vSection)));
	if (V_BSTR(&vSection)[0])
	{
		wcscat(wSubKey,L"\\");
		wcscat(wSubKey,V_BSTR(&vSection));
	}
	VariantClear(&vSection);
	lResult = RegOpenKeyEx(HKEY_CURRENT_USER, wSubKey, 0, KEY_READ, &hKey);
	if(lResult != ERROR_FILE_NOT_FOUND)
		lResult = ERROR_SUCCESS;
	else if (lResult == ERROR_SUCCESS)
	{
	VARIANT vKey;
		VariantInit(&vKey);
		VariantChangeType(&vKey,&Key,0,VT_BSTR);
		lResult = RegDeleteKey(hKey, V_BSTR(&vKey));
		VariantClear(&vKey);
	}
	RegCloseKey(hKey);
	if(lResult != ERROR_SUCCESS)
		diag(ERR_INVALID_FUNCTION_VALUE);
}

VARIANT WINAPI VBA_GetAllSettings(BSTR AppName,BSTR Section)
{
BSTR wSubKey;
HKEY hKey;
LONG lResult;
SAFEARRAY *psa;
SAFEARRAYBOUND saBound;
VARIANT vResult;

	V_VT(&vResult) = VT_NULL;
	wSubKey = SysAllocStringLen(AppName,SysStringLen(AppName)+1+SysStringLen(Section));
	if (Section[0])
	{
		wcscat(wSubKey,L"\\");
		wcscat(wSubKey,Section);
	}
	lResult = RegOpenKeyEx(HKEY_CURRENT_USER, wSubKey, 0, KEY_READ, &hKey);
	SysFreeString(wSubKey);
	if (lResult == ERROR_SUCCESS)
	{
	BYTE Data[1024]; /* what is max reg entry? */
	DWORD Type;
	DWORD cbData = sizeof(Data); /* yes, in bytes */
	VARIANT v;
	long l[2];
	WCHAR *Key;
Key = NULL;/* unimplemented - need to enumerate values !!!!!!! */
		saBound.cElements = 1;
		saBound.lLbound = 0;
		psa = SafeArrayCreate(VT_VARIANT,2,&saBound);
		V_VT(&vResult) = VT_VARIANT | VT_ARRAY;
		V_ARRAY(&vResult) = psa;
		lResult = RegQueryValueEx(hKey,Key,NULL,&Type,Data,&cbData);
		if(lResult == ERROR_SUCCESS)
			switch (Type)
			{
			case REG_DWORD:
				V_VT(&v) = VT_I4;
				V_I4(&v) = *(ULONG *)Data;
				break;
			case REG_SZ:
				V_VT(&v) = VT_BSTR;
				V_BSTR(&v) = SysAllocString((LPOLESTR)Data); /* REG_SZ is null terminated UNICODE */
				break;
			default:
/* need to convert more REG types to BSTR */
				diag(ERR_INVALID_FUNCTION_VALUE);
			}
		l[0] = 0;
		l[1] = 0;
		SafeArrayPutElement(psa,l,Key); /* put key */
		l[1] = 1;
		SafeArrayPutElement(psa,l,&v); /* put Variant containing data */
	}
	RegCloseKey(hKey);
 	if(lResult != ERROR_SUCCESS)
		diag(ERR_INVALID_FUNCTION_VALUE);
	return vResult;
}

VARIANT WINAPI VBA_CallByName(IDispatch * Object,BSTR ProcName,VBA_VbCallType CallType,VARIANT* Args)
{
long LBound,UBound;
HRESULT hr;
DISPID dispID[1];
DISPPARAMS dispParams;
VARIANT vResult;
EXCEPINFO ExcepInfo;
UINT ArgErr;
VARIANTARG rgvarg[1]; /* what should this be? */

	if (!(V_VT(Args) & VT_ARRAY))
		diag(ERR_INVALID_FUNCTION_VALUE); /* expecting array */
	if (V_VT(Args) != (VT_VARIANT | VT_ARRAY))
		diag(ERR_INVALID_FUNCTION_VALUE); /* expecting Variant array */
	if(SafeArrayGetDim(V_ARRAY(Args)) != 1)
		diag(ERR_INVALID_FUNCTION_VALUE); /* invalid number of dimensions */
	memset(&dispParams,0,sizeof(dispParams));
	SafeArrayGetLBound(V_ARRAY(Args),1,&LBound);
	SafeArrayGetUBound(V_ARRAY(Args),1,&UBound);
	dispParams.cArgs = UBound-LBound+1;
	dispParams.rgvarg = rgvarg;
	hr = IDispatch_GetIDsOfNames(Object,&IID_NULL,&ProcName,1,0,dispID);
	if (FAILED(hr))
		diag(ERR_INVALID_FUNCTION_VALUE);
#ifdef NEVER
long l;
	for(l=LBound;l<=UBound;l++)
	{
		VARIANT vBool;
		SafeArrayGetElement(V_ARRAY(Args),&l,&vBool);
		l++;
		if (V_BOOL(&vBool))
		{
			SafeArrayGetElement(V_ARRAY(Args),&l,&vResult);
			return vResult;
		}
	}
	hr = Object->Invoke(dispID,&IID_NULL,0,CallType,DispParams,&vResult,&ExcepInfo,&ArgErr);
#else
	SafeArrayLock(V_ARRAY(Args));
	SafeArrayPtrOfIndex(V_ARRAY(Args),&LBound,&dispParams);
	hr = IDispatch_Invoke(Object,dispID[1],&IID_NULL,0,(WORD)CallType,&dispParams,&vResult,&ExcepInfo,&ArgErr);
	SafeArrayUnlock(V_ARRAY(Args));
#endif
	if (FAILED(hr))
		diag(ERR_INVALID_FUNCTION_VALUE);
	return vResult;
}

/* Declaration: Math  0 . 0  */
VARIANT WINAPI VBA_Abs(VARIANT* Number)
{
VARIANT vResult;
	VariantInit(&vResult);
	if (VarAbs(Number,&vResult))
		abort();
	return vResult;
}

double WINAPI VBA_Atn(double Number)
{
	return atan(Number);
}

double WINAPI VBA_Cos(double Number)
{
#ifdef __GNUC__ /* don't know proper way to avoid inline error message */
double cos(double d);
#endif
	return cos(Number);
}

double WINAPI VBA_Exp(double Number)
{
	return exp(Number);
}

double WINAPI VBA_Log(double Number)
{
	return log10(Number);
}

void WINAPI VBA_Randomize(VARIANT* Number)
{
VARIANT vNumber;
	VariantInit(&vNumber);
	VariantChangeType(&vNumber,Number,0,VT_R8);
	if (V_R8(&vNumber) == 0.)
		srand((UINT)(time(NULL)+1)); /* +1 in case time is 0 */
	else
		srand((UINT)((DOUBLE)rand()/(DOUBLE)RAND_MAX*V_R8(&vNumber)));
/*	VariantClear(&vNumber);*/ /* not needed for VT_R8 */
}

float WINAPI VBA_Rnd(VARIANT* Number)
{
VARIANT vNumber;
static float last_rand = 1.; /* 1. cannot be generated */
float f;

	if (V_VT(Number) == VT_EMPTY)
		f = 1.;
	else
		{
		VariantInit(&vNumber);
		VariantChangeType(&vNumber,Number,0,VT_R4);
		f = V_R4(&vNumber);
/*		VariantClear(&vNumber);*/ /* not needed for VT_R4 */
		}
	if (f < 0.)
		srand((UINT)f);
	if (f != 0. || last_rand == 1.)
		last_rand = (float)rand()/(float)RAND_MAX;
	return last_rand;
}

double WINAPI VBA_Sin(double Number)
{
#ifdef __GNUC__ /* don't know proper way to avoid inline error message */
double sin(double d);
#endif
	return sin(Number);
}

double WINAPI VBA_Sqr(double Number)
{
#ifdef __GNUC__ /* don't know proper way to avoid inline error message */
double sqrt(double d);
#endif
	return sqrt(Number);
}

double WINAPI VBA_Tan(double Number)
{
	return tan(Number);
}

VARIANT WINAPI VBA_Sgn(VARIANT* Number)
{
VARIANT vNumber,vResult;
	VariantChangeType(&vNumber,Number,0,VT_R8);
	V_VT(&vResult) = VT_I2;
	V_I2(&vResult) = V_R8(&vNumber) > 0. ? 1 : V_R8(&vNumber) < 0. ? -1 : 0;
	return vResult;
}

VARIANT WINAPI VBA_Round(VARIANT* Number,long NumDigitsAfterDecimal)
{
VARIANT vResult;
	VariantInit(&vResult);
	if (VarRound(Number,NumDigitsAfterDecimal,&vResult))
		abort();
	return vResult;
}

/* Declaration: Financial  0 . 0  */
double WINAPI VBA_SLN(double Cost,double Salvage,double Life)
{
	diag(ERR_NOT_IMPLEMENTED);
}

double WINAPI VBA_SYD(double Cost,double Salvage,double Life,double Period)
{
	diag(ERR_NOT_IMPLEMENTED);
}

double WINAPI VBA_DDB(double Cost,double Salvage,double Life,double Period,VARIANT* Factor)
{
	diag(ERR_NOT_IMPLEMENTED);
}

double WINAPI VBA_IPmt(double Rate,double Per,double NPer,double PV,VARIANT* FV,VARIANT* Due)
{
	diag(ERR_NOT_IMPLEMENTED);
}

double WINAPI VBA_PPmt(double Rate,double Per,double NPer,double PV,VARIANT* FV,VARIANT* Due)
{
	diag(ERR_NOT_IMPLEMENTED);
}

double WINAPI VBA_Pmt(double Rate,double NPer,double PV,VARIANT* FV,VARIANT* Due)
{
	diag(ERR_NOT_IMPLEMENTED);
}

double WINAPI VBA_PV(double Rate,double NPer,double Pmt,VARIANT* FV,VARIANT* Due)
{
	diag(ERR_NOT_IMPLEMENTED);
}

double WINAPI VBA_FV(double Rate,double NPer,double Pmt,VARIANT* PV,VARIANT* Due)
{
	diag(ERR_NOT_IMPLEMENTED);
}

double WINAPI VBA_NPer(double Rate,double Pmt,double PV,VARIANT* FV,VARIANT* Due)
{
	diag(ERR_NOT_IMPLEMENTED);
}

double WINAPI VBA_Rate(double NPer,double Pmt,double PV,VARIANT* FV,VARIANT* Due,VARIANT* Guess)
{
	diag(ERR_NOT_IMPLEMENTED);
}

double WINAPI VBA_IRR(double* ValueArray,VARIANT* Guess)
{
	diag(ERR_NOT_IMPLEMENTED);
}

double WINAPI VBA_MIRR(double* ValueArray,double FinanceRate,double ReinvestRate)
{
	diag(ERR_NOT_IMPLEMENTED);
}

double WINAPI VBA_NPV(double Rate,double* ValueArray)
{
	diag(ERR_NOT_IMPLEMENTED);
}


/* Declaration: _HiddenModule  0 . 0  */
VARIANT WINAPI VBA_Array(VARIANT* ArgList)
{
	diag(ERR_NOT_IMPLEMENTED);
}

BSTR WINAPI VBA__B_str_InputB(long Number,short FileNumber)
{
#ifdef NEVER
FILBLK *fb;
BSTR bstr;
TEXT *buf;

	fb = inlfi(FileNumber,NULL);
	buf = sw_memalloc(Number+1,"input buffer",0);
	inlun(fb,buf,Number);
	buf[Number] = 0;
	inlwu(fb);
	bstr = sw_SysAllocStringLen((BSTR)buf,(Number+1)/2); /* is this ok for byte data? */
	sw_memfree(buf,Number+1,"input buffer");
	return bstr;
#else
	diag(ERR_NOT_IMPLEMENTED);
#endif
}

VARIANT WINAPI VBA__B_var_InputB(long Number,short FileNumber)
{
VARIANT vResult;
	V_VT(&vResult) = VT_BSTR;
	V_BSTR(&vResult) = VBA__B_str_InputB(Number,FileNumber);
	return vResult;
}

BSTR WINAPI VBA__B_str_Input(long Number,short FileNumber)
{
#ifdef NEVER
FILBLK *fb;
BSTR bstr;
TEXT *buf;

	fb = inlfi(FileNumber,NULL);
	buf = sw_memalloc(Number+1,"input buffer",0);
	inlun(fb,buf,Number);
	buf[Number] = 0;
	inlwu(fb);
	bstr = sw_SysAllocStringLen(NULL,Number); /* is this ok for byte data? */
	sw_MultiByteToWideChar(CP_ACP,MB_PRECOMPOSED,buf,size+1,bstr,size+1);
	sw_memfree(buf,Number+1,"input buffer");
	return bstr;
#else
	diag(ERR_NOT_IMPLEMENTED);
#endif
}

VARIANT WINAPI VBA__B_var_Input(long Number,short FileNumber)
{
VARIANT vResult;
	V_VT(&vResult) = VT_BSTR;
	V_BSTR(&vResult) = VBA__B_str_Input(Number,FileNumber);
	return vResult;
}

void WINAPI VBA_Width(short FileNumber,short Width)
{
	diag(ERR_NOT_IMPLEMENTED);
}

long WINAPI VBA_VarPtr(void* Ptr)
{
	return (long)Ptr;
}

long WINAPI VBA_StrPtr(BSTR Ptr)
{
	return (long)Ptr;
}

long WINAPI VBA_ObjPtr(IUnknown * Ptr)
{
	return (long)Ptr;
}
