
#include "vbtcom.h"
#include "test3/c/stdole_i.h"
#include "test3/c/VBA_i.h"
#include "vba.c"

#define FuncBody { FailedHR(0x12345678,L"Not implemented"); }

/* Declaration: Strings  0 . 0  */
short WINAPI VBA_Asc(BSTR String) FuncBody
BSTR WINAPI VBA__B_str_Chr(long CharCode) FuncBody
VARIANT WINAPI VBA__B_var_Chr(long CharCode) FuncBody
BSTR WINAPI VBA__B_str_LCase(BSTR String) FuncBody
VARIANT WINAPI VBA__B_var_LCase(VARIANT* String) FuncBody
BSTR WINAPI VBA__B_str_Mid(BSTR String,long Start,VARIANT* Length) FuncBody
VARIANT WINAPI VBA__B_var_Mid(VARIANT* String,long Start,VARIANT* Length) FuncBody
BSTR WINAPI VBA__B_str_MidB(BSTR String,long Start,VARIANT* Length) FuncBody
VARIANT WINAPI VBA__B_var_MidB(VARIANT* String,long Start,VARIANT* Length) FuncBody
VARIANT WINAPI VBA_InStr(VARIANT* Start,VARIANT* String1,VARIANT* String2,VBA_VbCompareMethod Compare) FuncBody
VARIANT WINAPI VBA_InStrB(VARIANT* Start,VARIANT* String1,VARIANT* String2,VBA_VbCompareMethod Compare) FuncBody
BSTR WINAPI VBA__B_str_Left(BSTR String,long Length) FuncBody
VARIANT WINAPI VBA__B_var_Left(VARIANT* String,long Length) FuncBody
BSTR WINAPI VBA__B_str_LeftB(BSTR String,long Length) FuncBody
VARIANT WINAPI VBA__B_var_LeftB(VARIANT* String,long Length) FuncBody
BSTR WINAPI VBA__B_str_LTrim(BSTR String) FuncBody
VARIANT WINAPI VBA__B_var_LTrim(VARIANT* String) FuncBody
BSTR WINAPI VBA__B_str_RightB(BSTR String,long Length) FuncBody
VARIANT WINAPI VBA__B_var_RightB(VARIANT* String,long Length) FuncBody
BSTR WINAPI VBA__B_str_Right(BSTR String,long Length) FuncBody
VARIANT WINAPI VBA__B_var_Right(VARIANT* String,long Length) FuncBody
BSTR WINAPI VBA__B_str_RTrim(BSTR String) FuncBody
VARIANT WINAPI VBA__B_var_RTrim(VARIANT* String) FuncBody
BSTR WINAPI VBA__B_str_Space(long Number) FuncBody
VARIANT WINAPI VBA__B_var_Space(long Number) FuncBody
VARIANT WINAPI VBA__B_var_StrConv(VARIANT* String,VBA_VbStrConv Conversion,long LocaleID) FuncBody
BSTR WINAPI VBA__B_str_String(long Number,VARIANT* Character) FuncBody
VARIANT WINAPI VBA__B_var_String(long Number,VARIANT* Character) FuncBody
BSTR WINAPI VBA__B_str_Trim(BSTR String) FuncBody
VARIANT WINAPI VBA__B_var_Trim(VARIANT* String) FuncBody
BSTR WINAPI VBA__B_str_UCase(BSTR String) FuncBody
VARIANT WINAPI VBA__B_var_UCase(VARIANT* String) FuncBody
VARIANT WINAPI VBA_StrComp(VARIANT* String1,VARIANT* String2,VBA_VbCompareMethod Compare) FuncBody
BSTR WINAPI VBA__B_str_Format(VARIANT* Expression,VARIANT* Format,VBA_VbDayOfWeek FirstDayOfWeek,VBA_VbFirstWeekOfYear FirstWeekOfYear) FuncBody
VARIANT WINAPI VBA__B_var_Format(VARIANT* Expression,VARIANT* Format,VBA_VbDayOfWeek FirstDayOfWeek,VBA_VbFirstWeekOfYear FirstWeekOfYear) FuncBody
VARIANT WINAPI VBA_Len(VARIANT* Expression) FuncBody
VARIANT WINAPI VBA_LenB(VARIANT* Expression) FuncBody
unsigned char WINAPI VBA_AscB(BSTR String) FuncBody
BSTR WINAPI VBA__B_str_ChrB(unsigned char CharCode) FuncBody
VARIANT WINAPI VBA__B_var_ChrB(unsigned char CharCode) FuncBody
short WINAPI VBA_AscW(BSTR String) FuncBody
BSTR WINAPI VBA__B_str_ChrW(long CharCode) FuncBody
VARIANT WINAPI VBA__B_var_ChrW(long CharCode) FuncBody
BSTR WINAPI VBA_FormatDateTime(VARIANT* Expression,VBA_VbDateTimeFormat NamedFormat) FuncBody
BSTR WINAPI VBA_FormatNumber(VARIANT* Expression,INT NumDigitsAfterDecimal,VBA_VbTriState IncludeLeadingDigit,VBA_VbTriState UseParensForNegativeNumbers,VBA_VbTriState GroupDigits) FuncBody
BSTR WINAPI VBA_FormatPercent(VARIANT* Expression,INT NumDigitsAfterDecimal,VBA_VbTriState IncludeLeadingDigit,VBA_VbTriState UseParensForNegativeNumbers,VBA_VbTriState GroupDigits) FuncBody
BSTR WINAPI VBA_FormatCurrency(VARIANT* Expression,INT NumDigitsAfterDecimal,VBA_VbTriState IncludeLeadingDigit,VBA_VbTriState UseParensForNegativeNumbers,VBA_VbTriState GroupDigits) FuncBody
BSTR WINAPI VBA_WeekdayName(INT Weekday,short Abbreviate,VBA_VbDayOfWeek FirstDayOfWeek) FuncBody
BSTR WINAPI VBA_MonthName(INT Month,short Abbreviate) FuncBody
BSTR WINAPI VBA_Replace(BSTR Expression,BSTR Find,BSTR Replace,long Start,long Count,VBA_VbCompareMethod Compare) FuncBody
BSTR WINAPI VBA_StrReverse(BSTR Expression) FuncBody
BSTR WINAPI VBA_Join(VARIANT* SourceArray,VARIANT* Delimiter) FuncBody
VARIANT WINAPI VBA_Filter(VARIANT* SourceArray,BSTR Match,short Include,VBA_VbCompareMethod Compare) FuncBody
long WINAPI VBA_InStrRev(BSTR StringCheck,BSTR StringMatch,long Start,VBA_VbCompareMethod Compare) FuncBody
VARIANT WINAPI VBA_Split(BSTR Expression,VARIANT* Delimiter,long Limit,VBA_VbCompareMethod Compare) FuncBody

/* Declaration: Conversion  0 . 0  */
BSTR WINAPI VBA__B_str_Hex(VARIANT* Number) FuncBody
VARIANT WINAPI VBA__B_var_Hex(VARIANT* Number) FuncBody
BSTR WINAPI VBA__B_str_Oct(VARIANT* Number) FuncBody
VARIANT WINAPI VBA__B_var_Oct(VARIANT* Number) FuncBody
long WINAPI VBA_MacID(BSTR Constant) FuncBody
BSTR WINAPI VBA__B_str_Str(VARIANT* Number) FuncBody
VARIANT WINAPI VBA__B_var_Str(VARIANT* Number) FuncBody
double WINAPI VBA_Val(BSTR String) FuncBody
BSTR WINAPI VBA_CStr(VARIANT* Expression) FuncBody
unsigned char WINAPI VBA_CByte(VARIANT* Expression) FuncBody
short WINAPI VBA_CBool(VARIANT* Expression) FuncBody
CY WINAPI VBA_CCur(VARIANT* Expression) FuncBody
DATE WINAPI VBA_CDate(VARIANT* Expression) FuncBody
VARIANT WINAPI VBA_CVDate(VARIANT* Expression) FuncBody
short WINAPI VBA_CInt(VARIANT* Expression) FuncBody
long WINAPI VBA_CLng(VARIANT* Expression) FuncBody
float WINAPI VBA_CSng(VARIANT* Expression) FuncBody
/*double WINAPI VBA_CDbl(VARIANT* Expression) FuncBody*/
VARIANT WINAPI VBA_CVar(VARIANT* Expression) FuncBody
VARIANT WINAPI VBA_CVErr(VARIANT* Expression) FuncBody
BSTR WINAPI VBA__B_str_Error(VARIANT* ErrorNumber) FuncBody
VARIANT WINAPI VBA__B_var_Error(VARIANT* ErrorNumber) FuncBody
VARIANT WINAPI VBA_Fix(VARIANT* Number) FuncBody
VARIANT WINAPI VBA_Int(VARIANT* Number) FuncBody
VARIANT WINAPI VBA_CDec(VARIANT* Expression) FuncBody

/* Declaration: FileSystem  0 . 0  */
void WINAPI VBA_ChDir(BSTR Path) FuncBody
void WINAPI VBA_ChDrive(BSTR Drive) FuncBody
short WINAPI VBA_EOF(short FileNumber) FuncBody
long WINAPI VBA_FileAttr(short FileNumber,short ReturnType) FuncBody
void WINAPI VBA_FileCopy(BSTR Source,BSTR Destination) FuncBody
VARIANT WINAPI VBA_FileDateTime(BSTR PathName) FuncBody
long WINAPI VBA_FileLen(BSTR PathName) FuncBody
VBA_VbFileAttribute WINAPI VBA_GetAttr(BSTR PathName) FuncBody
void WINAPI VBA_Kill(VARIANT* PathName) FuncBody
long WINAPI VBA_Loc(short FileNumber) FuncBody
long WINAPI VBA_LOF(short FileNumber) FuncBody
void WINAPI VBA_MkDir(BSTR Path) FuncBody
void WINAPI VBA_Reset() FuncBody
void WINAPI VBA_RmDir(BSTR Path) FuncBody
long WINAPI VBA_Seek(short FileNumber) FuncBody
void WINAPI VBA_SetAttr(BSTR PathName,VBA_VbFileAttribute Attributes) FuncBody
BSTR WINAPI VBA__B_str_CurDir(VARIANT* Drive) FuncBody
VARIANT WINAPI VBA__B_var_CurDir(VARIANT* Drive) FuncBody
short WINAPI VBA_FreeFile(VARIANT* RangeNumber) FuncBody
BSTR WINAPI VBA_Dir(VARIANT* PathName,VBA_VbFileAttribute Attributes) FuncBody

/* Declaration: DateTime  0 . 0  */
VARIANT WINAPI VBA__B_var_DateGet() FuncBody
void WINAPI VBA__B_str_DateLet(BSTR putval) FuncBody
void WINAPI VBA__B_var_DateLet(VARIANT putval) FuncBody
BSTR WINAPI VBA__B_str_DateGet() FuncBody
VARIANT WINAPI VBA_DateSerial(short Year,short Month,short Day) FuncBody
VARIANT WINAPI VBA_DateValue(BSTR Date) FuncBody
VARIANT WINAPI VBA_Day(VARIANT* Date) FuncBody
VARIANT WINAPI VBA_Hour(VARIANT* Time) FuncBody
VARIANT WINAPI VBA_Minute(VARIANT* Time) FuncBody
VARIANT WINAPI VBA_Month(VARIANT* Date) FuncBody
VARIANT WINAPI VBA_NowGet() FuncBody
VARIANT WINAPI VBA_Second(VARIANT* Time) FuncBody
VARIANT WINAPI VBA__B_var_TimeGet() FuncBody
void WINAPI VBA__B_str_TimeLet(BSTR putval) FuncBody
void WINAPI VBA__B_var_TimeLet(VARIANT putval) FuncBody
BSTR WINAPI VBA__B_str_TimeGet() FuncBody
/*float WINAPI VBA_TimerGet() FuncBody*/
VARIANT WINAPI VBA_TimeSerial(short Hour,short Minute,short Second) FuncBody
VARIANT WINAPI VBA_TimeValue(BSTR Time) FuncBody
VARIANT WINAPI VBA_Weekday(VARIANT* Date,VBA_VbDayOfWeek FirstDayOfWeek) FuncBody
VARIANT WINAPI VBA_Year(VARIANT* Date) FuncBody
VARIANT WINAPI VBA_DateAdd(BSTR Interval,double Number,VARIANT* Date) FuncBody
VARIANT WINAPI VBA_DateDiff(BSTR Interval,VARIANT* Date1,VARIANT* Date2,VBA_VbDayOfWeek FirstDayOfWeek,VBA_VbFirstWeekOfYear FirstWeekOfYear) FuncBody
VARIANT WINAPI VBA_DatePart(BSTR Interval,VARIANT* Date,VBA_VbDayOfWeek FirstDayOfWeek,VBA_VbFirstWeekOfYear FirstWeekOfYear) FuncBody
VBA_VbCalendar WINAPI VBA_CalendarGet() FuncBody
void WINAPI VBA_CalendarLet(VBA_VbCalendar putval) FuncBody

/* Declaration: Information  0 . 0  */
long WINAPI VBA_Erl() FuncBody
VBA__ErrObject WINAPI VBA_Err() FuncBody
VBA_VbIMEStatus WINAPI VBA_IMEStatus() FuncBody
short WINAPI VBA_IsArray(VARIANT* VarName) FuncBody
short WINAPI VBA_IsDate(VARIANT* Expression) FuncBody
short WINAPI VBA_IsEmpty(VARIANT* Expression) FuncBody
short WINAPI VBA_IsError(VARIANT* Expression) FuncBody
short WINAPI VBA_IsMissing(VARIANT* ArgName) FuncBody
short WINAPI VBA_IsNull(VARIANT* Expression) FuncBody
short WINAPI VBA_IsNumeric(VARIANT* Expression) FuncBody
short WINAPI VBA_IsObject(VARIANT* Expression) FuncBody
BSTR WINAPI VBA_TypeName(VARIANT* VarName) FuncBody
VBA_VbVarType WINAPI VBA_VarType(VARIANT* VarName) FuncBody
long WINAPI VBA_QBColor(short Color) FuncBody
long WINAPI VBA_RGB(short Red,short Green,short Blue) FuncBody

/* Declaration: Interaction  0 . 0  */
void WINAPI VBA_AppActivate(VARIANT* Title,VARIANT* Wait) FuncBody
void WINAPI VBA_Beep() FuncBody
/*VARIANT WINAPI VBA_CreateObject(BSTR Class,BSTR ServerName) FuncBody*/
short WINAPI VBA_DoEvents() FuncBody
VARIANT WINAPI VBA_GetObject(VARIANT* PathName,VARIANT* Class) FuncBody
BSTR WINAPI VBA_InputBox(VARIANT* Prompt,VARIANT* Title,VARIANT* Default,VARIANT* XPos,VARIANT* YPos,VARIANT* HelpFile,VARIANT* Context) FuncBody
BSTR WINAPI VBA_MacScript(BSTR Script) FuncBody
/*VBA_VbMsgBoxResult WINAPI VBA_MsgBox(VARIANT* Prompt,VBA_VbMsgBoxStyle Buttons,VARIANT* Title,VARIANT* HelpFile,VARIANT* Context) FuncBody*/
void WINAPI VBA_SendKeys(BSTR String,VARIANT* Wait) FuncBody
double WINAPI VBA_Shell(VARIANT* PathName,VBA_VbAppWinStyle WindowStyle) FuncBody
VARIANT WINAPI VBA_Partition(VARIANT* Number,VARIANT* Start,VARIANT* Stop,VARIANT* Interval) FuncBody
VARIANT WINAPI VBA_Choose(float Index,VARIANT* Choice) FuncBody
VARIANT WINAPI VBA__B_var_Environ(VARIANT* Expression) FuncBody
BSTR WINAPI VBA__B_str_Environ(VARIANT* Expression) FuncBody
VARIANT WINAPI VBA_Switch(VARIANT* VarExpr) FuncBody
VARIANT WINAPI VBA__B_var_Command() FuncBody
BSTR WINAPI VBA__B_str_Command() FuncBody
VARIANT WINAPI VBA_IIf(VARIANT* Expression,VARIANT* TruePart,VARIANT* FalsePart) FuncBody
BSTR WINAPI VBA_GetSetting(BSTR AppName,BSTR Section,BSTR Key,VARIANT Default) FuncBody
void WINAPI VBA_SaveSetting(BSTR AppName,BSTR Section,BSTR Key,BSTR Setting) FuncBody
void WINAPI VBA_DeleteSetting(BSTR AppName,VARIANT Section,VARIANT Key) FuncBody
VARIANT WINAPI VBA_GetAllSettings(BSTR AppName,BSTR Section) FuncBody
VARIANT WINAPI VBA_CallByName(IDispatch * Object,BSTR ProcName,VBA_VbCallType CallType,VARIANT* Args) FuncBody

/* Declaration: Math  0 . 0  */
VARIANT WINAPI VBA_Abs(VARIANT* Number) FuncBody
double WINAPI VBA_Atn(double Number) FuncBody
double WINAPI VBA_Cos(double Number) FuncBody
double WINAPI VBA_Exp(double Number) FuncBody
double WINAPI VBA_Log(double Number) FuncBody
void WINAPI VBA_Randomize(VARIANT* Number) FuncBody
float WINAPI VBA_Rnd(VARIANT* Number) FuncBody
double WINAPI VBA_Sin(double Number) FuncBody
double WINAPI VBA_Sqr(double Number) FuncBody
double WINAPI VBA_Tan(double Number) FuncBody
VARIANT WINAPI VBA_Sgn(VARIANT* Number) FuncBody
VARIANT WINAPI VBA_Round(VARIANT* Number,long NumDigitsAfterDecimal) FuncBody

/* Declaration: Financial  0 . 0  */
double WINAPI VBA_SLN(double Cost,double Salvage,double Life) FuncBody
double WINAPI VBA_SYD(double Cost,double Salvage,double Life,double Period) FuncBody
double WINAPI VBA_DDB(double Cost,double Salvage,double Life,double Period,VARIANT* Factor) FuncBody
double WINAPI VBA_IPmt(double Rate,double Per,double NPer,double PV,VARIANT* FV,VARIANT* Due) FuncBody
double WINAPI VBA_PPmt(double Rate,double Per,double NPer,double PV,VARIANT* FV,VARIANT* Due) FuncBody
double WINAPI VBA_Pmt(double Rate,double NPer,double PV,VARIANT* FV,VARIANT* Due) FuncBody
double WINAPI VBA_PV(double Rate,double NPer,double Pmt,VARIANT* FV,VARIANT* Due) FuncBody
double WINAPI VBA_FV(double Rate,double NPer,double Pmt,VARIANT* PV,VARIANT* Due) FuncBody
double WINAPI VBA_NPer(double Rate,double Pmt,double PV,VARIANT* FV,VARIANT* Due) FuncBody
double WINAPI VBA_Rate(double NPer,double Pmt,double PV,VARIANT* FV,VARIANT* Due,VARIANT* Guess) FuncBody
double WINAPI VBA_IRR(double* ValueArray,VARIANT* Guess) FuncBody
double WINAPI VBA_MIRR(double* ValueArray,double FinanceRate,double ReinvestRate) FuncBody
double WINAPI VBA_NPV(double Rate,double* ValueArray) FuncBody

/* Declaration: _HiddenModule  0 . 0  */
VARIANT WINAPI VBA_Array(VARIANT* ArgList) FuncBody
BSTR WINAPI VBA__B_str_InputB(long Number,short FileNumber) FuncBody
VARIANT WINAPI VBA__B_var_InputB(long Number,short FileNumber) FuncBody
BSTR WINAPI VBA__B_str_Input(long Number,short FileNumber) FuncBody
VARIANT WINAPI VBA__B_var_Input(long Number,short FileNumber) FuncBody
void WINAPI VBA_Width(short FileNumber,short Width) FuncBody
long WINAPI VBA_VarPtr(void* Ptr) FuncBody
long WINAPI VBA_StrPtr(BSTR Ptr) FuncBody
long WINAPI VBA_ObjPtr(IUnknown * Ptr) FuncBody
