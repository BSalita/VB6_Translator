Attribute VB_Name = "vbtdefs"
Option Explicit

' Priority cleanup:
' collections instead of .item, .add, ...
' invalid .exe path, display better message
' implement error processing, show offending line, proc, module name
' document overview, routines
' lookup routines vars, declares, ...
' Implement "Make Project Group" menu item.
' Don't replace Make menu items, insert VBT versions?
' change progress bar title to .vbp filename.

' Reasons for non-compiling perfectly
' 1. VB compile errors. If it doesn't compile in VB, it can't get to the VBT step. Try compiling without vbt addin.
' 2. Unreferenced labels - common. This is the only acceptable compiler warning (MSVC).
' 3. VB compiler errors - rare
' 4. Unimplemented Automation data types or data type conversions
' 5. Exe or DLL output filepath is invalid. UnAddIn vbt and try compiling DLL/Exe.
' 6. Another instance of VB is open which is using vbt add-in. (possible conflicts with simultaneous uses).
' 7. Invalid Project typelib reference or component.
' 8. Name conflict with emitter - struct, typedef, default etc. - audiol1a\project1.vbp
' 9. Use of non-English characters in symbols - Sub SalvaAlterações() in autoba1a\autobak.vbp
' 10. use of [string] for variable names may produce invalid emitter output. ggetvi1r\ptransl_.vbp
' 11. Invalid filenames such as Aux, Con, Nul, etc.
' 12. Invalid file path specified in .vbp file (path for loading source files or PATH32).
' 13. TypeLib has invalid interface name for emitter (GIF89LibCtl.Gif89a s/b GIF89LibCtl_Gif89a)
' 14. Required default member not in typelib interface
' 15. Unsupported type conversions, such as StrTo{numeric}, occuring on bad parameter conversions, a% = "", KeyAscii = vbTab, etc.
' 16. Invalid typelib data such as bad parameter data type in access.dbengine (dao350.dll)
' 17. Invalid syntax not caught by VB - Error.Raise 1 s/b Err.Raise 1
' 18. MSVC LngOr(l,0x80000000) gives signed/unsigned mismatch for unknown reasons (others don't)(maybe its the signed extension of 0x80000000)
' 19. VBT has a bug in processing UserControl module type
' 20. Name conflicts uncaught by VB such as attempting to define "Line" in a form or "Name","Print" (its a statement).
' 21. Opening .vbp instead of .vbg
' 22. Declare statement must not have duplicate parameter names
' 23. Missing or newer typelib is selected in project references.
' 24. Event name must not later be used to define a procedure name. RaiseEvent will fail.
' 25. Previous compiler error may prevent rerunning of compiler. Remove VBT addin and any vbt.exe processes, then re-addin vbt.
' 26. VBT has bugs in processing Designer, UserControl and PropertyPage forms. DBExample\DBExample.vbp
' 27. Errors occuring in VB's automated conversion from older VB versions to VB6 may confuse VBT. Try reloading converted project and try VBT again. (rare)
' 28. Error in cemitter was caused by bad vbt.exe in memory, probably unreleased from previous use because of error.
' 29. Unsupported conversions (e.g. unsigned longs) or erroneous conversions (e.g. string-to-double caused by use of + instead of &)
' 30. Public variables used as For statement counters are not supported.
' 31. Comments are discarded.
' 32. Numeric constants may not retain same representation as original. Hex, floating point precision differences.

' General
' 1. comb through sources and update this list with to do(s).
' 2. redo debug messages - provide debug level

' Compiler - clean up
' 1. tokstring s/b tokvariant
' 2. implement individual tokens, not aggragates
' 3. implement error handling instead of err.raise 1, use paramarray, error levels, (no) return,error severity
' 4. finish storing module, proc obj ref in varmodule, etc
' 6. create collections of symbol references for xref
' 7. create dictionary classes and change Collection to Dictionary where appropriate
' 8. For all operators, do sanity check that's in OprIs, OprNeg and oprOperateUnary routine
' 10. Implement SourceInfo in vbToken and elsewhere
' 11. Should vbNest use vbLabel or sourceInfo?
' 12. create collection of nesting collections for use as sanity checking in EndProc?
' 13. Need to trap overflow (cint) in vbTokenize
' 14. How to properly lookup components in tlib
' 15. Implement passing of ByVal keyword in declareArgs to emitter
' 16. implement class/form init/terminate - always generate in same place in vtable
' 18. Implement BEGIN/END/VERSION (don't really have to)
' 19. raise error if ParamArray is used with Optional keyword
' 21. Disallow changing of vbToken fields (except for tokoutput) after vbt processing. Can't let any emitter change any member.
' 22. Check for case sensitive compares
' 23. check for missing type coercions in compiler
' 24. Implement Implements compile time checking, and runtime support
' 27. Disallow string * length when not permitted.
' 29. Resolve confusion about object defaults - objExcel.Range("A3") = "new total"
' 30. d = i + i or d = i / i didn't generate double divide?
' 31. Dim x as Variant / ReDim x(1) as Integer works - why? implement and test
' 33. Insert With statements
' 34. emit comments? emit source lines?
' 35. Make all TLI Class lookups use ClassInterfaceFromTLib
' 36. Each symbol must have InvokeKind, to be And'ed or compared with usage (Get,Let)
' 37. Generate prototypes of all possible operator and operand combinations (auto generate BoolNot() etc macros)
' 40. Namespace conflicts with emitters - audiol1a\project1.vbp
' 41. Fix LSet, RSet
' 43. Implement WithValue processing of control arrays
' 45. suppress outputing of unreferenced labels (MSVC warns)
' 46. clear up confusion between dtInterfaceInfo, tokInterfaceInfo and ReturnTypes
' 47. implement test that verifies data types after constant fold (OprNeg?)
' 48. default values for tli functions improperly implemented - shouldn't be emitting them - function handles them - need to insert default value into table
' 49. Use dictionary key or implement token for named argument dispatch objects - o.mysub parm1 := 123
' 50. Implement Group compile
' 51. Implement control array hidden interface members UBound(kw), LBound(kw), Count
' 52. Get TLib Enum's (tokConstInfo) to output TypeDef
' 53. Use tokEnumMember instead of Variant?
' 54. Change 1 to FUNCFLAG_FRESTRICTED
' 55. Process VBG files.
' 56. Implement Cancel button
' 57. Selectively turn on/off debugging
' 58. Fix dsp emitter for new project files.
' 59. Output typelibs to {GUID} directory. Don't regenerate if exists and timedate is OK. Regenerate if refemitter is updated?
' 60. Check DLL/Exe output filename for validity before compiling.
' 61. Append interesting errors to log file.
' 62. Allow multiple instances of vbt add-in (in several active VB apps).
' 63. Why doesn't amazin1a\project1.vbp find msflxgrd.ocx control?
' 64. Is there a problem with licensed controls? Can't open tlib? msflxgrd.ocx
' 65. Doesn't handle other alphabets - Sub SalvaAlterações() in autoba1a\autobak.vbp
' 66. Currency code generation in C emitter needs testing. Use CY variables and expressions, ByRefs, funcs, etc.
' 67. What to do about [string] syntax that becomes invalid variable names?
' 68. Need to make vbVariable same as vbDataType to better support UDT typedefs.
' 69. In AddIn, should tokenize once, let each pass process tokens as needed.
' 70. Trap invalid filenames such as Aux, Con, Nul, etc.
' 71. Option Base/Compare statements aren't implemented in emitters.
' 72. implement checking for Enums and Types having no members
' 73. fix case when passing a typed object to a ByRef Object, does sender or receiver (or both) do a IDispatchQI/typeQI?
' 74. ForUDT1.vbp - Public For statement counters not supported.
' 75. Redo attribute DEFAULT setting PRIVATE or PUBLIC as appropriate.
' 76. Write function to do bit tests instead of using Not flag ...
' 77. Can IsObject() always be used instead of Not o As Nothing? If not, write MyIsObject?
' 78. Save constant orignal constant representation (hex, decimal point, float fraction)?

' Compiler - test
' 1. test event processing
' 2. Enum lookup in local and module and global scopes
' 3. Implement enum name prefix - myenum.myvalue. What are scope rules in classes?
' 4. Empty parameters ",,"
' 5. Missing parameters should be tokMissing, not Empty variant
' 6. Test a project that contains all component types (enum vbext_ComponentType)
' 7. Test multiple vbt Add-ins
' 8. Test Public (TypeLib) Enums, Structs (VB doesn't support Unions) -- Typedefs?
' 9. Test LSet, RSet with ByRef
' 10. Test each type of statement followed by ":" (multi-statement line)
' 11. Notable test programs:
'       to do: collect other mentions of .vbp files and add to this list.
'       csdtpckr\csdtpckr.vbp - tests Const scoping, UserControl.xxx, more
'       floodp1r\paladin.vbp - tests ActiveX Designer
'       get_dr1r\tv.vbp - name conflict with HWND_NOTOPMOST in winuser.h. HWND_NOTOPMOST is mistakenly undefined in form1.c
'       getfil2r\datasmoo.vbp - array assignment
'       multgrid\gridsamp.vbp - typelib SAFEARRAY (dao_i.h)
'       realse1a\webpager.vbp - preprocessor, StrToByteArray
'       demooutlook\projet1.vbp [sic] - default member missing in interface
'       CryptoAPI\CryptoAPI_Demo.vbp - SafeArray ByRefs
'       multicolsort\multicolsort.vbp
'       karaoke\demokar.vbp - RaiseEvent
'       pguide\workshop\pamdi\pa_mdi.vbp - default member of public variable
'       udl2\udlfile.vbp - ByRef parameter missing ByRef flag - PromptEdit
'       pbaspedit\pbaspediter.vbp - user control bug - function not found rtfmain.SetFocus
'       urlhist\urlHistory.vbp - bad tlb file - don't know why?
'       grid\start_me.vbg - ocx compile - aborts in functionargs
'       varSpace\varSpace.vbp - fails because Event and Friend have same name. Extensibility gets confused.
'       FormShaper\SharperTest.vbp - fails with bad reference typelib
'       PlayerX\XPlayer.vbp - Designer bug?
'       superwrite\project1.vbp - won't compile because of non-ascii characters.
'       compttt\prottt.vbp - Won't compile, default problem with Data control.
'       VBT has a bug in processing Designer properties. DBExample\DBExample.vbp
'       DropdownCalculator has user control bug
'       prjDataEnv\prjSchool.prj - designer form bug
'       aeoncode\rpgame.vbp - usercontrol bug - simple control
'       dynarep\prjdynarep.vbp - designer bug
'       DiskDataInfo\dskDataGroup.vbg - usercontrol tlib error
'       QNavigator\QNavigator.vbp - usercontrol.line - line method not found
'       EasterEgg\EasterEgg.vbp - usercontrol .vbp not accessable from other .vbp
'       ForUDT1.vbp - Public For statement counters not supported.
'       MthVwPrj.vbp - VB - won't load frmMthVw.frm
'       ADODataSortControl.vbp - doesn't compile - usercontrol
'       prjparam.vbp - designer
'       prjshape.vbp - designer
'       logo\logo.vbp - uses o.line
' 12. Do any functions, other than Input and InputB, allow # before FileNumber?
' 13. Unresolved bugs: ActiveXDesigner .Show derivation. helper1a\treeview.vbp
' 14. What other array assigns are allowed besides StrToByteArray, xxxArrayToVar?
' 15. More ReDim testing needed. Variant and ReDim "As" not fully tested. ReDimVariant1.vbp
' 16. Const accepts multiple defs. Which other statement do (#Const, DefType)?
' 17. Test : and Else at end of line for all statements
' 18. When passing Any, should address of variable be passed? What about Variant (contents or address)?
' 19. timebilling.vbp gets DLL error, never investigated
' 20. What does New on ReDim really do? Are all elements New? ReDim c(1) as New Class1
' 21. Create project(s) that selects all typelib references.
' 22. Create project(s) that use every available component.

' Documentation
' 1. comment and pretty code
' 2. create tables of stuff in html (funcs, operators, keywords, statements, etc)
' 3. create table of pcodes
' 4. xyz.badmember = xyz.badmember ' note that rhs gives error first
' 5. Recommend using Option Explicit
' 6. Can't compile may be caused by invalid syntax or invalid .EXE name. Remove vbt Addin, check .EXE name.
' 7. Width statement withdrawn beginning with VB4.
' 8. Can't continue Enum members using "_"
' 9. Unlike others of the ilk, the following methods are documented to have Single arguments but they're actually Variants in the IDL: PaintPicture, Scale, ScaleX, ScaleY
' 10. Form1("Text1")(0) is actually Form1.Controls.Item("Text1")(0)
' 11. VB uses polymorphism - Left vs Left(s,1) is Form1.Left vs. VBA.Left
' 12. Suggest Option Explicit always be used.
' 13. Invalid filenames such as Aux, Con, Nul, etc.
' 14. ReDim accepts more than just simple variable names. They can be any array expression.
' 15. Const accepts multiple defs. Const x=1,y=2,z=3
' 16. Multiple definition on a line presents line numbering problems to IDE extensibility
' 17. ElseIf 1 Then i = 1 is valid in VB6 (may be VB6 bug?)
' 18. MsgBox a$ + a$ becomes a$ & a$ but MsgBox a$ + 1 becomes MsgBox CStr(CDbl(a$)+1)
' 19. Invalid syntax not caught by VB - Error.Raise 1 s/b Err.Raise 1
' 20. When passing a typed object to a ByRef Object, does sender or receiver (or both) do a IDispatchQI/typeQI?
' 21. !a = !a is legal (With o;!a=!a;End With}
' 22. "Dim Line" should be caught by VB as invalid in a form.
' 23. VB allows Declare statement to have duplicate parameter names
' 24. I set process priority on vbt.exe to below normal so it doesn't hog CPU time.
' 25. Use of Option Explicit statement is highly recommended.
' 26. Don't use statement names as variables (Dim Name), even if VB allows it.
' 27. Event name must not later be used to define a procedure name. RaiseEvent will fail.
' 28. What is a VB statement really? Who defines it?
' 29. Must have a .vbp file.
' 30. ASP not supported.
' 31. "Time = 0" Time is a propput with a ReturnType of VARIANT * (instead of last parameter). Seems wrong.
' 32. Exception case: non-Optional last parameter must follow optional parameters for Prop Let/Set.
' 33. "Print" not allowed as method name. VB allows "Print" syntax but is it a special Print method or standard method, don't use it.
' 34. Having problem with finding MSCommLib (mscomm32.ocx SP4) (and a few others). Downloaded mscomm.zip and reinstalled. Needed to reboot. Perhaps regsvr32 worked but I forgot to reboot???
' 35. When VBT has problem with typelib fullpath, object browser (F2) doesn't display library file path too.
' 36. interfaces of typekind = dispatch do not generate default members in Let statements. Form1.Picture = LoadPicture("...") - no defaults
' 37. Labels names seem a bit arbitrary. Can't name labels with statement names (For, Let, any member in VBA tlib) but can use local variable names, or methods and properties within same module. What's up with that?
' 38. symbols allowed as labels (but perhaps should be reserved words)
'       Error
'       Line
'       Mid
'       MidB
'       Name
'       Object
'       Property
'       Step
'       Time
'       Width
'     VBA members not definable or referenceable as labels
'       String,Len,LenB,CStr,CByte,CBool,CDate,CInt,CLng,CSng,CDbl,CVar,CVerr,Fix,Int,Seek,Abs,Sgn,Array,InputB,Input
'     VBA members definable as label but not referenceable in GoTo
'       CDec,Date,DoEvents
'       So what is "Date:Date:Date", "CDec:CDec:CDec" "CDec(1):..."?
' 39. Contrary to documentation, ReDim allows "As New Class"
' 40. "Set col.Item(1) = Nothing" inserts a default member on LHS.
' 41. Contrary to documentation, For statement counter may be an array element or even a UDT member.
' 42. Public For statement counters are special cases. What about Privates? Are they being handled correctly? Can't use address-of variables. Inefficient.
' 43. Const s as string * 1 = 1 = 1 = 1 - What is the length? One.
' 44. Don't optimize expressions if using VBEmitter. Things like flag Oring will become unintelligable constants.
' 45. VB syntax is so haphazard that there's no definitive syntactic description.
' 46. Why does "Call DoEvents" or "Call Abs(1)" create syntax errors but "Call Beep" doesn't?
' 47. What is datatype of result of 2 / 2? Answer: Double.
' 48. Step shouldn't be a permitted variable (array) name, ambiguous - For i = 1 to 10 Step(1)
' 49. CreateObject probably returns IUnknown - Set o = CreateObject() - Set earlyBound = CreateObject() - CreateObject() is Nothing
' 50. Default interface is a Class concept needed by VB because VB only can reference the default (or first) interface, never a 2nd interface if first was default

' Questions for experts
' 1. How to enumerate component libs?
' 2. Do non-public methods have GUIDs and vtables? Do they CoCreateInstance or just class factory?
' 3. How to get groupname?
' 4. How to get at hidden typelib interfaces? Are there form control collections? (form1.controls)
' 5. Can Local (On Local Error) be safely ignored.
' 6. Why isn't IsExternalType flag set for some TypeInfos describing external libraries? MsHTML, MSsHTMLCtl
' 7. How does OutLook's Items.Item(1).Subject --> Debug.Print Items(1) - Item and Subject are not defaults but are somehow invoked.
' 8. When passing a typed object to a ByRef Object, does sender or receiver (or both) do a IDispatchQI/typeQI?
' 9. Why do some typelibs not have a FullPath?
' 10. "Time = 0" Time is a propput with a ReturnType of VARIANT * (instead of last parameter). Seems wrong.
' 11. Why don't interfaces of typekind = dispatch generate default members in Let statements? Form1.Picture = LoadPicture("...") - no defaults
' 12. Why do large VB programs which use "Compile to Native Code" option hang or take so long to compile? Recompiling using PCode complete within seconds.
' 13. Why confusion about VBA labels (see 38 above)

' API Wishes
' 1. Want API to do VarTypeToCDataType conversion.

' VB Wishes
' 1. TypeLib: implement tags in object
' 2. VBIDE: implment compile-without-executing menu option
' 3. VBIDE: persist maximize window upon reload of project
' 4. VBIDE: better add-in debugging
' 5. VBADDIN: make group name available in Add-in extensibility model
' 6. VBADDIN: allow enumeration of Project->Component selected and whole list
' 7. VBADDIN: provide easier way of accessing typelibs of controls (Textbox)
' 8. TypeLib: direct access to default members by invokekind
' 9. VBIDE: should be able to do "Current Project" Find command even if all codepanes closed
' 10. TypeLib: implement Parent in all applicable classes (MemberInfo)
' 11. TypeLib: Easier VTable view
' 12. VB RT: allow CreateObject to accept GUID -- CreateObject("{...}",...)
' 13. Allow interface specification -- Dim i as New TLib.ClassName.InterfaceName
' 14. VBIDE: Want browser or intellisense to display enum values
' 15. Want option in New Projects dialog to create an empty bas module.
' 16. Require For statement counters to be local scope. More efficient, easier for compiler writers ;-).

' VB Bugs
' 1. VB compiler: catch "1" & o is nothing as syntax error
' 2. VB compiler: Property Let/Set allows assignment to property function name - should be synatx error
' 3. TypeLib: Some property puts have paramcount = 0?
' 4. VB compiler: allows "Dim mid(1,2,3) as Integer" "Mid(s,1,2) = mid(s,1,2)" - Where Mid is func and mid is array
' 5. AddIn currentComponent.CodeModule.Members collection duplicates names for "Event sub1" and "Sub sub1".
' 6. Width statement documented but won't compile as of VB4.
' 7. Can't continue Enum members using "_"
' 8. Enum declares can change case u/l according to case of last reference
' 9. Erroneous error message - Public Const = 1 gives message that says user-defined types can't be public - they can be
' 10. Intellisense not working - TypeOf v Is tagMyType - UDTs not displayed
' 11. Next a, b - contrary to syntax doc
' 12. Duplicate params allowed in declare statement - Declare Sub DuplicateParams Lib "lib" (same as integer,same as integer)
' 13. Error = 1 - Assignment isn't documented. Keyword isn't colorized. Results in Err.Raise 424!??
' 14. VB6 IDE editor doesn't highlight CDec()
' 15. v.print 1;2 - should be invalid. Otherwise, what would argument of ; look like?
' 16. Allowing statement after ElseIf Then (e.g. ElseIf 1 then i = 0)
' 17. "Dim Error : On Error GoTo 10" - VB should flag "Dim Error" as invalid definition.
' 18. Invalid syntax not caught by VB - Error.Raise 1 s/b Err.Raise 1
' 19. "Dim Line" should be invalid in a form.
' 20. Declare statement should not allow multiply defined parameter names
' 21. Exit Function within a Property Get is not flaged as syntax error.
' 22. "Dim Name" shouldn't be allowed, Name is a statement.
' 23. Extensibility models fails when Event and method have same name. e.g. "Event SomeName()" "Sub SomeName()". Sub will obscure Event causing RaiseEvent to fail.
' 24. See list of symbols allowed as label names. In particular, Object is allowed but all other types aren't. Mid, MidB, Error are suspect
' 25. Contrary to documentation, ReDim allows "As New Class"
' 26. IDE gave a #else statement as the line number content of a Property Get
' 27. Error: and Error.Raise should be flagged at syntax time.
' 28. Sub name obscured by local variable - Sub s: Dim s: s (error - thinks s is string, not recursive call)
' 29. Labels confusion. See 38 above.
' 30. Implements can't handle method names containing underscores. e.g. Open_Routine

' Other compilers
' 1. implement AlphaBasic source parsing
' 2. Output initing code for object variables declared with New.

' C Run-time Support
' 1. comb through vbt.h - code for all data types
' 2. implement functions defined in vbt.h

' Emitter
' 1. splitup vbtsubs.c into init routines, subs, rtl support
' 2. implement vb subs
' 3. C - Finish implementing TypeOf - how does VB do it?
' 4. C - output casts in C emitter
' 5. C - always output stdole_i.h (and other component libs)?
' 6. C - implement non-zero lower bound array (output adjust macro?)
' 7. VBView - needs work - implement property variables
' 8. C - Forms doesn't support control collection, need controlCollection interfaces
' 9. C - Hmmm. Forms controls must be QI-able through Form This ptr.
' 10. C - implement safearrays
' 11. C - is array bound info generated?
' 12. C - integral mis-match warning occasional generated
' 13. implement debug output of class variables and UDTs.
' 14. How to get group name?
' 15. Which path should emitter files be output to?
' 16. C - Rethink when to trigger remaking of reference files (.c .h).
' 17. C - rethink naming conventions. eliminate windows.h? when to use "_" prefix?
' 18. C - make one routine emit all variable names, then can check for namespace conflicts - audiol1a\project1.vbp
' 19. Refs - Output dependent _i.h files (such as office_i.h in vbide_i.h)
' 20. VBA_CDec - return value VarByRef(VBA_CDec()) not handled correctly
' 21. C - finish Decimal and Currency types
' 22. C - Need to finish implementing WithValue(0) - needs casting for datatype of UDT (WithValue(0)->item).
' 23. C - Implement pointer levels in cemitter
' 24. C - Resolve variable-name/C-keyword conflicts - default, continue - audiol1a\project1.vbp
' 25. VB - Don't optimize constant expressions, named bit flags are replaced by numeric constants.
' 26. VB - Me takes on Form name. Unload Me becomes Unload form1
' 27. VB - Ugly Step - For Date = 11 to 12 becomes #1/10/1900# To #1/11/1900# Step #12/31/1899#
' 28. VB - Output Method's Dim statements in same place as declared, not always at top of method
' 29. VB - Recompose Single-Line-If as single-line instead of block-If.
' 30. VB - Create option to allow/disallow output of implicit variables declarations.

' Other Emitters
' 0. Create ref to .h class and precompiled .lib project
' 1. Output .VBR file.
' 2. binary pcode emitter (in C?)
' 3. C pcode emitter (one func per pcode)
' 4. C++ emitter (std, ATL, MFC)
' 5. java emitter
' 6. Corba DII instead of OLE Automation
' 7. decompiler
' 9. output VBP file?
' 10. SOAP
' 11. CDL
' 12. XML
' 13. validation emitter, checks object model
' 14. VB statistics
' 15. TransMeta Crusoe
' 16. Lint
'       conversions
'       unref labels/vars/enums/tlibs/params ...
'       xref
'       default objects
'       variant
'       check enum param values against enums
'       constant expressions against allowed range (subscripts)
' 17. PocketPC
' 18. BNF
' 19. C#
' 20. Delphi
' 21. Objective C
' 22. obfuscator

' fixme: make these into tokCP, tokOP, tokComma
Public Enum SpecialTypes
    SPECIAL_EMPTY
'    special_eq
    SPECIAL_CP
    SPECIAL_OP
    special_comma
'    special_period
    SPECIAL_NS ' used in file I/O statements
    special_sc
'    special_underline
'    special_colon
End Enum

Public Enum OptimizeFlags
    OptimizeNone = &H0
    OptimizeConstantExpressions = &H1
    ProcessCompilerExpression = OptimizeConstantExpressions Or &H2
End Enum

Public Type ApplicationData
    appName As String
End Type

Public Type SourceInfo
    sourceFile As String
    sourceLine As Long
End Type

Public cVariantTypes As New Collection
Public appData As ApplicationData
Public gOptimizeFlag As Long
Public cCompilerDirectives As New Collection
Public cDataTypes As New Collection
Public cKeywords As New Collection
Public cUnaryOperators As New Collection
Public cBinaryOperators As New Collection
Public cSpecialTypes As New Collection
Public cStatements As New Collection
Public g_cPredefinedConstant As New Collection
Public g_main_inited As Boolean

Public group As New vbGroup
'Public GlobalVariables As New Collection
Public currentProject As vbPrj
Public currentModule As vbModule
Public currentProc As proctable
'Public msTLI As TypeLibInfo
'Public msTLIDeclarations As tli.Declarations ' MSVBVM60 type declarations
'Public msTLIConstants As tli.Constants ' MSVBVM60 type declarations
Public oRPN As New vbRPNize ' fixme: use .bas instead
Public CompilerIfStack As New Collection
Public currentComponent As VBComponent
Public currentLineNumber As Long
'Public lines() As String
'Public cmlines() As String
Public vbOprAdd As New vbOprAdd ' need to put this in each opr class
Public oMe As Object ' Need new mneumonic
Public tliVB As reference ' Needed for Form coclass
Public tlibVBA As TypeLibInfo ' Needed for label checking
Public UnknownInterfaceInfo As InterfaceInfo
Public DispatchInterfaceInfo As InterfaceInfo
Public OutputPath As String
Public Scope As New vbtScope ' fixme: use .bas instead
Public ProcessDeclares As Boolean
Public PassNumber As Integer
Public LastQualifier As vbToken ' last symbol
Public AssignmentOperator As vbToken ' last symbol before assignment operator
Public oTok As New vbTokenize ' fixme: use .bas instead

Private oElse As New vbStmtElse
Private oElseIf As New vbStmtElseIf
Private oIf As New vbStmtIf
Private oLet As New vbStmtLet
Private oSLM As New vbStmtLocalMethod
Private EmptyCollection As New Collection

Sub Main()
'TLIApplication.ResolveAliases = False

Dim token As New vbToken

cDataTypes.Add New vbDataType, "ANY"
cDataTypes.Item(cDataTypes.Count).dtDataType = VT_VOID ' Any is valid within Declare statements
cDataTypes.Add New vbDataType, "BOOLEAN"
cDataTypes.Item(cDataTypes.Count).dtDataType = vbBoolean
cDataTypes.Add New vbDataType, "BYTE"
cDataTypes.Item(cDataTypes.Count).dtDataType = vbByte
cDataTypes.Add New vbDataType, "CURRENCY"
cDataTypes.Item(cDataTypes.Count).dtDataType = vbCurrency
cDataTypes.Add New vbDataType, "DATE"
cDataTypes.Item(cDataTypes.Count).dtDataType = vbDate
cDataTypes.Add New vbDataType, "DECIMAL"
cDataTypes.Item(cDataTypes.Count).dtDataType = vbDecimal
cDataTypes.Add New vbDataType, "DOUBLE"
cDataTypes.Item(cDataTypes.Count).dtDataType = vbDouble
cDataTypes.Add New vbDataType, "INTEGER"
cDataTypes.Item(cDataTypes.Count).dtDataType = vbInteger
cDataTypes.Add New vbDataType, "LONG"
cDataTypes.Item(cDataTypes.Count).dtDataType = vbLong
cDataTypes.Add New vbDataType, "OBJECT"
cDataTypes.Item(cDataTypes.Count).dtDataType = vbObject
cDataTypes.Add New vbDataType, "SINGLE"
cDataTypes.Item(cDataTypes.Count).dtDataType = vbSingle
cDataTypes.Add New vbDataType, "STRING"
cDataTypes.Item(cDataTypes.Count).dtDataType = vbString
cDataTypes.Add New vbDataType, "USERDEFINEDTYPES"
cDataTypes.Item(cDataTypes.Count).dtDataType = vbUserDefinedType
cDataTypes.Add New vbDataType, "VARIANT"
cDataTypes.Item(cDataTypes.Count).dtDataType = vbVariant

' collect VB data types for operand coercions
'' fixme: add more automation data types
'' fixme: need to test whether VB coerces to VB types or ignores (try unsigned + 1E11)
cVariantTypes.Add CBool(1), CStr(vbBoolean)
cVariantTypes.Add CByte(1), CStr(vbByte)
cVariantTypes.Add CCur(1), CStr(vbCurrency)
cVariantTypes.Add CDate(1), CStr(vbDate)
cVariantTypes.Add CDbl(1), CStr(vbDouble)
cVariantTypes.Add CDec(1), CStr(vbDecimal) ' VB IDE bug: Hey, cstr(vb6 doesn't highlight CDec.
cVariantTypes.Add CInt(1), CStr(vbInteger)
cVariantTypes.Add CLng(1), CStr(vbLong)
cVariantTypes.Add CSng(1), CStr(vbSingle)
cVariantTypes.Add CStr(1), CStr(vbString)
'cVariantTypes.Add Monochrome, CStr(VT_UI4) ' (UINT32)1
cVariantTypes.Add Nothing, CStr(vbObject)

cKeywords.Add Keywords.KW_ADDRESSOF, "ADDRESSOF"
cKeywords.Add Keywords.KW_AS, "AS"
cKeywords.Add Keywords.KW_BYREF, "BYREF"
cKeywords.Add Keywords.KW_BYVAL, "BYVAL"
cKeywords.Add Keywords.KW_DEBUG, "DEBUG"
cKeywords.Add Keywords.KW_DO, "DO"
cKeywords.Add Keywords.KW_EACH, "EACH"
cKeywords.Add Keywords.KW_ELSE, "ELSE"
cKeywords.Add Keywords.KW_ELSEIF, "ELSEIF"
cKeywords.Add Keywords.KW_FOR, "FOR"
cKeywords.Add Keywords.KW_FUNCTION, "FUNCTION"
cKeywords.Add Keywords.KW_GOTO, "GOTO"
cKeywords.Add Keywords.KW_GOSUB, "GOSUB"
cKeywords.Add Keywords.KW_IF, "IF"
cKeywords.Add Keywords.KW_IN, "IN"
' These shouldn't be keyword because they are valid expression terms
'cKeywords.Add Keywords.KW_INPUT, "INPUT" ' function!!!
'cKeywords.Add Keywords.KW_INPUTB, "INPUTB" ' function!!!
cKeywords.Add Keywords.KW_LBOUND, "LBOUND"
'cKeywords.Add Keywords.KW_LEN, "LEN"
'cKeywords.Add Keywords.KW_ME, "ME" - predefined variable
cKeywords.Add Keywords.KW_LOCAL, "LOCAL"
cKeywords.Add Keywords.KW_NEW, "NEW"
cKeywords.Add Keywords.KW_OPTIONAL, "OPTIONAL"
cKeywords.Add Keywords.KW_PARAMARRAY, "PARAMARRAY"
cKeywords.Add Keywords.KW_PRESERVE, "PRESERVE"
' fixme: this suggests that keyword table shouldn't be used. Test against char constant instead?
cKeywords.Add Keywords.KW_PROPERTY, "PROPERTY" ' not removed - valid as label name
cKeywords.Add Keywords.KW_SELECT, "SELECT"
cKeywords.Add Keywords.KW_STEP, "STEP" ' not removed - valid as label name
cKeywords.Add Keywords.KW_SUB, "SUB"
cKeywords.Add Keywords.KW_THEN, "THEN"
cKeywords.Add Keywords.KW_TO, "TO"
cKeywords.Add Keywords.KW_TYPE, "TYPE"
cKeywords.Add Keywords.KW_TYPEOF, "TYPEOF"
cKeywords.Add Keywords.KW_UBOUND, "UBOUND"
cKeywords.Add Keywords.KW_UNTIL, "UNTIL"
cKeywords.Add Keywords.KW_WHILE, "WHILE"
cKeywords.Add Keywords.KW_WITH, "WITH"
cKeywords.Add Keywords.KW_WITHEVENTS, "WITHEVENTS"

' some of these are dup vbOprxx calls and maybe should not use New
cUnaryOperators.Add New vbOprPos, "+" ' also add and catenate
cUnaryOperators.Add New vbOprNeg, "-" ' also binary subtraction
cUnaryOperators.Add New vbOprNot, "NOT"

cBinaryOperators.Add New vbOprCat, "&" ' also long var type, also radix convert
cBinaryOperators.Add New vbOprDiv, "/"
cBinaryOperators.Add New vbOprMul, "*"
cBinaryOperators.Add New vbOprAdd, "+" ' also catenate and unary positive
cBinaryOperators.Add New vbOprSub, "-" ' also unary negative
cBinaryOperators.Add New vbOprLT, "<"
cBinaryOperators.Add New vbOprLE, "<="
cBinaryOperators.Add New vbOprNE, "<>"
cBinaryOperators.Add New vbOprEQ, "="
cBinaryOperators.Add New vbOprGE, "=>"
cBinaryOperators.Add New vbOprGT, ">"
cBinaryOperators.Add New vbOprNE, "><"
cBinaryOperators.Add New vbOprGE, ">="
cBinaryOperators.Add New vbOprAnd, "AND"
cBinaryOperators.Add New vbOprEqv, "EQV"
cBinaryOperators.Add New vbOprImp, "IMP"
cBinaryOperators.Add New vbOprIs, "IS"
cBinaryOperators.Add New vbOprLike, "LIKE"
cBinaryOperators.Add New vbOprMod, "MOD"
cBinaryOperators.Add New vbOprOr, "OR"
cBinaryOperators.Add New vbOprXor, "XOR"
cBinaryOperators.Add New vbOprIDiv, "\"
cBinaryOperators.Add New vbOprPow, "^"

' fixme: make these into tokComma, etc.
cSpecialTypes.Add SpecialTypes.special_comma, ","
'cSpecialTypes.Add SpecialTypes.special_period, "."
cSpecialTypes.Add SpecialTypes.SPECIAL_CP, ")"
'cSpecialTypes.Add SpecialTypes.special_eq, "=" ' should this be here?
cSpecialTypes.Add SpecialTypes.SPECIAL_OP, "("
cSpecialTypes.Add SpecialTypes.SPECIAL_NS, "#" ' used in file I/O
cSpecialTypes.Add SpecialTypes.special_sc, ";" ' used in Print/Write/Debug statements
'cSpecialTypes.Add SpecialTypes.special_underline, "_"

cStatements.Add New vbStmtPrint, "?" ' want this?
cStatements.Add New vbStmtAttribute, "ATTRIBUTE" ' parsed, not fully implemented
cStatements.Add New vbStmtCall, "CALL" ' ByVal and ByRef keywords not implemented
cStatements.Add New vbStmtCase, "CASE"
cStatements.Add New vbStmtClose, "CLOSE"
cStatements.Add New vbStmtConst, "CONST" ' compiler directive
'cStatements.Add New vbStmtDate, "DATE"
cStatements.Add New vbStmtDebug, "DEBUG"
cStatements.Add New vbStmtDeclare, "DECLARE" ' compiler directive
cStatements.Add New vbStmtDef, "DEFBOOL" ' compiler directive
cStatements.Add New vbStmtDef, "DEFBYTE" ' compiler directive
cStatements.Add New vbStmtDef, "DEFCUR" ' compiler directive
cStatements.Add New vbStmtDef, "DEFDATE" ' compiler directive
cStatements.Add New vbStmtDef, "DEFDEC" ' compiler directive
cStatements.Add New vbStmtDef, "DEFDBL" ' compiler directive
cStatements.Add New vbStmtDef, "DEFINT" ' compiler directive
cStatements.Add New vbStmtDef, "DEFLNG" ' compiler directive
cStatements.Add New vbStmtDef, "DEFOBJ" ' compiler directive
cStatements.Add New vbStmtDef, "DEFSNG" ' compiler directive
cStatements.Add New vbStmtDef, "DEFSTR" ' compiler directive
cStatements.Add New vbStmtDef, "DEFVAR" ' compiler directive
cStatements.Add New vbStmtDim, "DIM"
cStatements.Add New vbStmtDo, "DO"
cStatements.Add New vbStmtElse, "ELSE"
cStatements.Add New vbStmtElseIf, "ELSEIF"
cStatements.Add New vbStmtEnd, "END"
'cStatements.Add New vbStmtEndIf, "ENDIF"
cStatements.Add New vbStmtEnum, "ENUM"
cStatements.Add New vbStmtErase, "ERASE"
cStatements.Add New vbStmtError, "ERROR" ' not removed - valid as label name
cStatements.Add New vbStmtEvent, "EVENT"
cStatements.Add New vbStmtExit, "EXIT"
cStatements.Add New vbStmtFriend, "FRIEND" ' not finished, need to implement private public static
cStatements.Add New vbStmtFor, "FOR"
cStatements.Add New vbStmtFunction, "FUNCTION"
cStatements.Add New vbStmtGet, "GET"
cStatements.Add New vbStmtPublic, "GLOBAL" ' deprecated, replaced by Public
cStatements.Add New vbStmtGoSub, "GOSUB"
cStatements.Add New vbStmtGoTo, "GOTO"
cStatements.Add New vbStmtIf, "IF"
cStatements.Add New vbStmtImplements, "IMPLEMENTS"
cStatements.Add New vbStmtInput, "INPUT"
cStatements.Add New vbStmtInput, "INPUTB" ' same as Input
cStatements.Add New vbStmtLet, "LET"
'cStatements.Add New vbStmtLineInput, "LINE"
cStatements.Add New vbStmtLock, "LOCK"
cStatements.Add New vbStmtLoop, "LOOP"
cStatements.Add New vbStmtLSet, "LSET"
cStatements.Add New vbStmtMid, "MID" ' not removed - valid as label name
cStatements.Add New vbStmtMidB, "MIDB" ' not removed - valid as label name
cStatements.Add New vbStmtName, "NAME" ' not removed - valid as label name
cStatements.Add New vbStmtNext, "NEXT"
cStatements.Add New vbStmtOn, "ON"
cStatements.Add New vbStmtOpen, "OPEN"
cStatements.Add New vbStmtOption, "OPTION" ' compiler directives
cStatements.Add New vbStmtPrint, "PRINT" ' need to implement , ; spc
cStatements.Add New vbStmtPrivate, "PRIVATE" ' not finished
cStatements.Add New vbStmtProperty, "PROPERTY" ' not removed - valid as label name
cStatements.Add New vbStmtPublic, "PUBLIC" ' not finished
cStatements.Add New vbStmtPut, "PUT"
cStatements.Add New vbStmtRaiseEvent, "RAISEEVENT"
cStatements.Add New vbStmtReDim, "REDIM"
cStatements.Add New vbStmtResume, "RESUME"
cStatements.Add New vbStmtReturn, "RETURN"
cStatements.Add New vbStmtRSet, "RSET"
cStatements.Add New vbStmtSeek, "SEEK"
cStatements.Add New vbStmtSelect, "SELECT"
cStatements.Add New vbStmtSet, "SET"
cStatements.Add New vbStmtStatic, "STATIC" ' not finished
cStatements.Add New vbStmtStop, "STOP"
cStatements.Add New vbStmtSub, "SUB"
'cStatements.Add New vbStmtTime, "TIME"
cStatements.Add New vbStmtType, "TYPE"
cStatements.Add New vbStmtUnLock, "UNLOCK"
cStatements.Add New vbStmtWend, "WEND"
cStatements.Add New vbStmtWhile, "WHILE"
' Width was removed from language as of VB4 although still documented
' Seems to compile (in VB6 SP3) so better leave it in.
' fixme: Width is only valid in Form modules?
'cStatements.Add New vbStmtWidth, "WIDTH"
cStatements.Add New vbStmtWith, "WITH"
cStatements.Add New vbStmtWrite, "WRITE" ' need to implement , ; spc

Set oElse = cStatements.Item("ELSE")
Set oElseIf = cStatements.Item("ELSEIF")
Set oIf = cStatements.Item("IF")
Set oLet = cStatements.Item("LET")

' Create predefined constants
' use class instead?
Set token = New vbToken
token.tokType = tokVariant
token.tokString = "WIN16"
token.tokValue = False
g_cPredefinedConstant.Add token, UCase(token.tokString)
Set token = New vbToken
token.tokType = tokVariant
token.tokString = "WIN32"
token.tokValue = True
g_cPredefinedConstant.Add token, UCase(token.tokString)

cCompilerDirectives.Add New vbCompilerDirectiveConst, "CONST"
cCompilerDirectives.Add New vbCompilerDirectiveIf, "IF"
cCompilerDirectives.Add New vbCompilerDirectiveElse, "ELSE"
cCompilerDirectives.Add New vbCompilerDirectiveElseIf, "ELSEIF"
cCompilerDirectives.Add New vbCompilerDirectiveEndIf, "END"

#If 0 Then
cCompilerDirectives.Add New vbCompilerDirectiveError, "ERROR" ' VB Extension
cCompilerDirectives.Add New vbCompilerDirectiveIfDef, "IFDEF" ' VB Extension
cCompilerDirectives.Add New vbCompilerDirectiveIfNDef, "IFNDEF" ' VB Extension
cCompilerDirectives.Add New vbCompilerDirectiveInclude, "INCLUDE" ' VB Extension
cCompilerDirectives.Add New vbCompilerDirectivePragma, "PRAGMA" ' VB Extension
#End If

g_main_inited = True

End Sub

Sub CompileVB()
Print #99, "CompileVB: dcl="; currentModule.dcllines
Dim tokens As Collection
currentLineNumber = 1
Dim dcllines As Long
dcllines = currentModule.dcllines
While currentLineNumber <= dcllines
    currentLineNumber = GetNextTokenizedLine(currentLineNumber, tokens)
    parseLine tokens
    If tokens.Count > 0 Then Err.Raise 1
Wend
End Sub

Sub parseLine(ByVal tokens As Collection)
Print #99, "ParseLine: t.c="; tokens.Count
Dim newTokens As New Collection
' Check for linenumber, label definition
Dim output_stack As Collection
If tokens.Count > 0 Then
    If tokens.Item(1).tokType = tokVariant Then
        Select Case varType(tokens.Item(1).tokValue)
        Case vbByte, vbInteger
            tokens.Item(1).tokValue = CLng(tokens.Item(1).tokValue)
        Case vbLong
            ' check some range?
        Case Else
            Err.Raise 1 ' Invalid line number
        End Select
        Set output_stack = New Collection
        AddLabelDef tokens.Item(1), output_stack
        tokens.Remove 1
    ElseIf tokens.Count > 1 Then ' label?
        Print #99, "Label: 1"
        If tokens.Item(1).tokType = toksymbol And tokens.Item(2).tokType = tokColon Then
            Print #99, "Label: 2"
            ' labels may not be a keyword or VBA Declarations member
' fixme: all this should be replaced by a collection lookup. info is static.
            Dim o As Object
            On Error Resume Next
            Set o = Nothing
            Set o = cStatements.Item(UCase(tokens.Item(1).tokString))
            If o Is Nothing Then Set o = cUnaryOperators.Item(UCase(tokens.Item(1).tokString))
            If o Is Nothing Then Set o = cBinaryOperators.Item(UCase(tokens.Item(1).tokString))
            If o Is Nothing Then Set o = cDataTypes.Item(UCase(tokens.Item(1).tokString))
            On Error GoTo 0
            Print #99, "Label: 3"
            If o Is Nothing Then
                Print #99, "Label: 4"
                If getKeyword(tokens.Item(1)) = 0 Then
                    Print #99, "Label: 5"
#If 1 Then
                    ' Using VB's compiler to catch invalid label identifiers.
                    ' Except the following identifiers are valid labels but invalid as label references (GoTo DoEvents).
                    If IsAny(UCase(tokens.Item(1).tokString), "CDEC", "DATE", "DOEVENTS") Then
                        Print #99, "Ambiguous label/method identifier " & """" & tokens.Item(1).tokString & """" & ". Use valid label name."
                        MsgBox "Ambiguous label/method identifier " & """" & tokens.Item(1).tokString & """" & ". Use valid label name."
                        Err.Raise 1
                    End If
                    Print #99, "Label: 7"
                    Set output_stack = New Collection
                    AddLabelDef tokens.Item(1), output_stack
                    tokens.Remove 1
                    tokens.Remove 1
#If 0 Then
                    Dim DeclarationInfo As DeclarationInfo
                    For Each DeclarationInfo In tlibVBA.Declarations
                        Print #99, "Label: 6 di="; DeclarationInfo.Name; " m.c="; DeclarationInfo.Members.Count
' Lookup by Item fails, strings not supported. Using For Each instead.
'                        On Error Resume Next
'                        Set o = DeclarationInfo.Members.Item(tokens.Item(1).tokString)
'                        On Error GoTo 0
'                        If Not o Is Nothing Then Exit For
                        Dim MemberInfo As MemberInfo
                        For Each MemberInfo In DeclarationInfo.Members
                            Print #99, MemberInfo.Name; ":"
                            If UCase(MemberInfo.Name) = UCase(tokens.Item(1).tokString) Then Exit For
                        Next
                        If Not MemberInfo Is Nothing Then Exit For
                    Next
                    Print #99, "Label: 7"
                    If DeclarationInfo Is Nothing Then
                        Print #99, "Label: 8"
                        Set output_stack = New Collection
                        AddLabelDef tokens.Item(1), output_stack
                        tokens.Remove 1
                        tokens.Remove 1
                    End If
#End If
#If 0 Then
                    Dim DeclarationInfo As DeclarationInfo
                    For Each DeclarationInfo In tlibVBA.Declarations
                        Print #99, "' "; tlibVBA.Name; "."; DeclarationInfo.Name
                        Dim MemberInfo As MemberInfo
                        For Each MemberInfo In DeclarationInfo.Members
                            Dim ss As String
                            If Left(MemberInfo.Name, 7) = "_B_var_" Or Left(MemberInfo.Name, 7) = "_B_str_" Then ss = Mid(MemberInfo.Name, 7 + 1) Else ss = MemberInfo.Name
                            Print #99, ss; ":"
                            Print #99, "goto "; ss
                        Next
                    Next
                    Err.Raise 1
#End If
#End If
                End If
            End If
        End If
    End If
End If

Print #99, "ParseLine: 2"
Do Until tokens.Count = 0
    ' remove leading colons
    Do While tokens.Item(1).tokType = tokColon
        tokens.Remove 1
        If tokens.Count = 0 Then GoTo 10
    Loop
Print #99, "ParseLine: 3"
    Dim s As vbStmt
    Dim Stmt As vbStmt
    Set LastQualifier = Nothing
    Set AssignmentOperator = Nothing
    On Error Resume Next
    Set Stmt = Nothing
    Set Stmt = cStatements.Item(UCase(tokens.Item(1).tokString))
    On Error GoTo 0
    Select Case UCase(tokens.Item(1).tokString)
    Case "LINE"
        If Not IsEOL(tokens) Then If UCase(tokens.Item(2).tokString) = "INPUT" Then Set Stmt = New vbStmtLineInput
    Case Else
        ' do nothing
    End Select
    If Stmt Is Nothing Then
Print #99, "ParseLine: 4 t="; tokens.Item(1).tokType
        If tokens.Item(1).tokType <> toksymbol And tokens.Item(1).tokType <> tokWith Then
            Print #99, "Invalid statement identifier """ & tokens.Item(1).tokString & """."
            MsgBox "Invalid statement identifier """ & tokens.Item(1).tokString & """."
            Err.Raise 1  ' Unrecognized statement
        End If
        If isAssignment(tokens, 1) Then
Print #99, "Stmt is an assignment"
            Set Stmt = oLet
            tokens.Add New vbToken, , 1 ' space holder
            tokens.Item(1).tokType = tokstatement
            tokens.Item(1).tokString = "Let"
        Else
Print #99, "Stmt is a method call"
            Set Stmt = oSLM ' vbStmtLocalMethod
        End If
    Else
        ' isAssignment determines LastQualifier, checks for invalid syntax
Print #99, "ParseLine: 5", UCase(tokens.Item(1).tokString)
        Select Case UCase(tokens.Item(1).tokString)
            Case "FOR", "LET", "SET", "RSET", "LSET"
                If Not isAssignment(tokens, 2) Then If UCase(tokens.Item(1).tokString) <> "FOR" Or UCase(tokens.Item(2).tokString) <> "EACH" Then Err.Raise 1
        End Select
        tokens.Item(1).tokType = tokstatement
    End If
Print #99, "ParseLine: 6 aq="; Not AssignmentOperator Is Nothing; " lq="; Not LastQualifier Is Nothing
    Stmt.stmtParse tokens
Print #99, "ParseLine: 7 tc="; tokens.Count
    If tokens.Count > 0 Then
Print #99, "ParseLine: 8 s="; tokens.Item(1).tokString
        If tokens.Item(1).tokType <> tokColon Then
            Print #99, "n.c="; currentProc.procNests.Count
            If currentProc.procNests.Count <> 0 Then
                Print #99, "bif.c="; currentProc.procBlockIfs.Count
                Print #99, "sif.c="; currentProc.procSingleIfs.Count
                If currentProc.procSingleIfs.Count <> 0 Then
Print #99, "1"
                    If Not currentProc.procNests.Item(1) Is currentProc.procSingleIfs.Item(1) Then MsgBox "Expecting EOL": Err.Raise 1 ' Expecting end of line - may be VB bug, doc error, or unimplemented syntax
Print #99, "2"
                    If UCase(tokens.Item(1).tokString) <> "ELSE" Then If (Not Stmt Is oIf) And (Not Stmt Is oElse) Then MsgBox "Expecting EOL": Err.Raise 1 ' Expecting end of line - may be VB bug, doc error, or unimplemented syntax
Print #99, "3"
                ElseIf currentProc.procBlockIfs.Count <> 0 Then
Print #99, "4"
                    If Not currentProc.procNests.Item(1) Is currentProc.procBlockIfs.Item(1) Then MsgBox "Expecting EOL": Err.Raise 1 ' Expecting end of line - may be VB bug, doc error, or unimplemented syntax
Print #99, "5"
                    If Not Stmt Is oElseIf Then MsgBox "Expecting EOL": Err.Raise 1 ' Expecting end of line - may be VB bug, doc error, or unimplemented syntax
Print #99, "6"
                Else
                    MsgBox "Expecting EOL"
                    Err.Raise 1  ' Expecting end of line - may be VB bug, doc error, unimplemented syntax
                End If
            Else
                MsgBox "Expecting EOL"
                Err.Raise 1  ' Expecting end of line - may be VB bug, doc error, unimplemented syntax
            End If
        End If
    End If
Loop

10
Print #99, "ParseLine: 10 cp="; Not currentProc Is Nothing
If Not currentProc Is Nothing Then
    Print #99, "ParseLine: 11"
    While currentProc.procSingleIfs.Count > 0
        Print #99, "ParseLine: 12"
        Dim SingleIfEndIf As vbToken
        If Not currentProc.procNests.Item(1) Is currentProc.procSingleIfs.Item(1) Then Err.Raise 1 ' currentproc.procnests.nestObj.stmtTok.tokstring invalid is single line If
        Set SingleIfEndIf = New vbToken
        SingleIfEndIf.tokType = tokstatement
        SingleIfEndIf.tokPCode = vbPCodeSingleIfEndIf
        SingleIfEndIf.tokString = "EndIf"
        Set output_stack = New Collection
        output_stack.Add SingleIfEndIf
        currentProc.procStatements.Add output_stack
        currentProc.procSingleIfs.Remove 1
        currentProc.procNests.Remove 1
    Wend
End If
Print #99, "ParseLine: done"

End Sub

Function isAssignment(ByVal tokens As Collection, ByVal first As Long) As Boolean
Dim i As Integer
Dim p As Integer
If tokens.Item(first).tokType = toksymbol Then Set LastQualifier = tokens.Item(first)
For i = first To tokens.Count - 1
    Select Case tokens.Item(i).tokString
        Case "("
            p = p + 1
        Case ")"
            p = p - 1
        Case Else
            If p = 0 Then
                Select Case tokens.Item(i).tokString
                    Case "."
                        If tokens.Item(i).tokType = tokWith Then
                            If i <> first Then Exit For
                        Else
                            i = i + 1
                            If tokens.Item(i).tokType <> toksymbol Then Err.Raise 1 ' Expecting qualifier
                        End If
                        Set LastQualifier = tokens.Item(i) ' moved here for "!a = !a"
                    Case "="
                        isAssignment = True
                        Set AssignmentOperator = tokens.Item(i)
                        Exit For
                    Case ":" ' multi-line statement
                        Exit For
                    Case Else
                        If i <> first Then Exit For
                End Select
            End If
    End Select
Next
If LastQualifier Is Nothing Then Print #99, "lq=nothing" Else Print #99, "lq="; LastQualifier.tokString
End Function

Function getTokenizedLine(ByVal ln As Long) As Collection
Print #99, "getTokenizedLine: ln="; ln

' fixme: This should grap already tokenized line?!
' But need to define cmtokens because line numbering (keys) are different than tokens
' Some tokenized lines can be used from saved collection, but not "Sub s:blah:End Sub"
' Initial keyword may be removed, just rescan line for now.
#If 1 Then
On Error Resume Next
Dim l As Long
l = currentModule.ilines.Item(CStr(ln))
Print #99, "ilines="; l
If l = 0 Then
    Set getTokenizedLine = New Collection
Else
    Dim s As String
    s = currentModule.lines.Item(CStr(l))
    Print #99, "lines="; s
    Set getTokenizedLine = oTok.TokenizeLine(s)
End If
On Error GoTo 0
#Else
Dim tokens As Collection
On Error Resume Next
Set tokens = New Collection
Set tokens = currentModule.tokens.Item(CStr(ln))
Dim ts As Collection ' fixme: make obsolete, just double check and correct for now
Set ts = New Collection
Set ts = oTok.TokenizeLine(currentModule.cmlines.Item(CStr(ln)))
On Error GoTo 0
If tokens.Count <> ts.Count Then
    Print #99, "getTokenizedLine: tokens.c=" & tokens.Count & "<> ts.c=" & ts.Count
'    MsgBox "getTokenizedLine: tokens.c=" & tokens.count & "<> ts.c=" & ts.count
    If tokens.Count <> 0 Then Err.Raise 1
    Set tokens = ts
End If
Dim token As vbToken
For Each token In tokens
    Print #99, token.tokCol; token.tokType; token.tokString
Next
Set getTokenizedLine = tokens
#End If
Print #99, "getTokenizedLine: gtl.c="; getTokenizedLine.Count
End Function

Function CheckVarDup(ByVal token As vbToken, variables As Collection) As Boolean
Dim v As vbVariable
On Error Resume Next
Set v = variables.Item(UCase(token.tokString))
If Not v Is Nothing Then
' Var dup can be either an invalid var dup or multiple vars on a Dim statement which AddIn delivers dup line numbers
    Print #99, "tc="; token.tokComponent Is Nothing; " vc="; v.varComponent Is Nothing
    Print #99, "tln="; token.tokLineNumber; " vln="; v.varLineNumber; " tcn="; token.tokComponent.Name; " vcn="; v.varComponent.Name
    ' if same linenumber and component, then ignore - due to AddIn oddity
' fixme - need to implement column number to avoid Dim A,A (but VB will catch anyway)
    If token.tokLineNumber <> v.varLineNumber Or Not token.tokComponent Is v.varComponent Then Err.Raise 1     ' dup variable def
    CheckVarDup = True
End If
End Function

Function getSpecialTypes(ByVal token As vbToken) As SpecialTypes
On Error Resume Next
getSpecialTypes = cSpecialTypes.Item(token.tokString)
End Function

Function getKeyword(ByVal token As vbToken) As Keywords
On Error Resume Next
getKeyword = cKeywords.Item(UCase(token.tokString))
End Function

' may be able to cut out some step in function processing
Sub getArrayVariableAddr(ByVal tokens As Collection, ByVal output_stack As Collection)
Print #99, "getArrayVariableAddr: tc="; tokens.Count; " osc="; output_stack.Count
Dim token As vbToken
' fixme: test that SymbolLookUp doesn't provide invalid array types
Set token = SymbolLookUp(gOptimizeFlag, tokens, output_stack, INVOKE_FUNC Or INVOKE_PROPERTYGET, False, True)
If token Is Nothing Then Err.Raise 1
Print #99, "getArrayVariableAddr: ts="; token.tokString; " dt="; token.tokDataType; " t="; token.tokType
'If token.tokType <> tokArrayVariable And token.tokType <> tokUDT And (token.tokType <> tokVariable Or token.tokDataType <> vbVariant) Then Err.Raise 1 ' must be a variable
' Could be LBound(v)
If token.tokType <> tokArrayVariable And token.tokType <> tokUDT And token.tokType <> tokVariantArgs And (token.tokType <> tokVariable Or (token.tokDataType And Not VT_BYREF) <> vbVariant) Then Err.Raise 1 ' must be a variable
' fixme: This brings up the point that VT_ARRAY may not be correct.
'   Should we use VT_VECTOR instead? What is difference?
If (token.tokDataType And Not VT_BYREF) <> vbVariant And Not CBool(token.tokDataType And (VT_ARRAY Or VT_VECTOR)) Then
    Print #99, "Expecting array name: " & token.tokString
    MsgBox "Expecting array name: " & token.tokString
    Err.Raise 1 ' Variable not defined
End If
' fixme: commented out because need to preserve tokUDT - any implications?
'token.tokType = tokArrayVariable
output_stack.Add token
Print #99, "getArrayVariableAddr: tc="; tokens.Count; " osc="; output_stack.Count
End Sub

Sub AddTLIDeclarationMember(ByVal di As DeclarationInfo, ByVal mi As MemberInfo)
Dim tlib As TypeLibInfo
Set tlib = di.Parent
Dim ref As vbtTLIRef
Set ref = Nothing
On Error Resume Next
Set ref = currentProject.prjTLIRefs.Item(tlib.Name)
On Error GoTo 0
Print #99, ref Is Nothing
If ref Is Nothing Then
Print #99, currentProject.prjTLIRefs.Count
    currentProject.prjTLIRefs.Add New vbtTLIRef, tlib.Name
Print #99, tlib.Name
    Set ref = currentProject.prjTLIRefs.Item(tlib.Name)
Print #99, ref Is Nothing
    Set ref.RefTLib = tlib
End If
Print #99, di.Name & "." & mi.Name
Dim dm As vbtTLIDeclarationMember
On Error Resume Next
Set dm = ref.RefDeclarationMembers.Item(di.Name & "." & mi.Name)
On Error GoTo 0
If dm Is Nothing Then
    ref.RefDeclarationMembers.Add New vbtTLIDeclarationMember, di.Name & "." & mi.Name
    Set dm = ref.RefDeclarationMembers.Item(di.Name & "." & mi.Name)
    Set dm.DeclarationInfo = di
    Set dm.MemberInfo = mi
End If
End Sub

Sub AddTLIInterfaceMember(ByVal ii As InterfaceInfo, ByVal mi As MemberInfo, ByVal rt As VarTypeInfo, ByVal ik As InvokeKinds)
Dim tlib As TypeLibInfo
On Error Resume Next
Print #99, "AddTLIInterfaceMember:"
Print #99, "ii="; Not ii Is Nothing
Print #99, "ii="; TypeInfoToVBType(ii)
Print #99, "ii.f="; ii.Parent.ContainingFile
Print #99, "mi.ik="; mi.InvokeKind; " ik="; ik
On Error GoTo 0
On Error Resume Next
Set tlib = TypeLibInfoFromFile(tliVB.FullPath) ' kludge: unable to get parent of controls (Textbox)
Set tlib = ii.Parent
On Error GoTo 0
Dim ref As vbtTLIRef
On Error Resume Next
Set ref = currentProject.prjTLIRefs.Item(tlib.Name)
On Error GoTo 0
Print #99, "ref="; ref Is Nothing
If ref Is Nothing Then
Print #99, currentProject.prjTLIRefs.Count
    currentProject.prjTLIRefs.Add New vbtTLIRef, tlib.Name
Print #99, tlib.Name
    Set ref = currentProject.prjTLIRefs.Item(tlib.Name)
Print #99, ref Is Nothing
    Set ref.RefTLib = tlib
End If
Dim s As String
s = ProcNameIK(ii.Name & "_" & mi.Name, IIf(mi.InvokeKind = 0, ik, mi.InvokeKind)) ' CStr(mi.InvokeKind)
Print #99, s
Dim im As vbtTLIInterfaceMember
On Error Resume Next
Set im = ref.RefInterfaceMembers.Item(s)
On Error GoTo 0
If im Is Nothing Then
    ref.RefInterfaceMembers.Add New vbtTLIInterfaceMember, s
    Set im = ref.RefInterfaceMembers.Item(s)
    Set im.InterfaceInfo = ii
    Set im.MemberInfo = mi
Print #99, "AddTLIInterfaceMember: mi="; mi.Name; " ik="; mi.InvokeKind; " pc="; mi.parameters.Count
    Set im.ReturnType = rt
    im.InvokeKind = ik
End If
End Sub

Sub ProcessArgs(ByRef rtii As InterfaceInfo, ByRef token As vbToken, ByVal tokens As Collection, ByVal arg_stack As Collection, ByVal output_stack As Collection, ByRef ik As InvokeKinds)
Print #99, "ProcessArgs: rtii="; rtii.Name; " ik="; ik
If arg_stack.Count = 0 Then
    Set rtii = Nothing
Else
    If token.tokVariable.varAttributes And VARIABLE_CONTROLARRAY Then
        If arg_stack.Count = 0 Then ik = INVOKE_FUNC Else If arg_stack.Count <> 1 Then Err.Raise 1
        AppendCollection arg_stack.Item(1), output_stack
        CoerceOperand gOptimizeFlag, output_stack, output_stack.Count, vbLong
        arg_stack.Remove 1
        If token.tokCount <> 0 Then Err.Raise 1
        token.tokCount = 1
    Else
        Print #99, "ProcessArgs: rtii=" & rtii.Name; " st="; token.tokPCodeSubType
        Dim mi As MemberInfo
    ' fixme: major kludge: pcodeSubType has 14 so using PROPERTYPUT but could be PUTREF
'' loops infinitely on "Form1.Controls(1).Value = 0"
    Do ' Need loop for oRset!CustomerId = "" (oRset.Fields.Item("CustomerId").Value = "")
        Dim ii As InterfaceInfo
        Set ii = rtii ' save for 2nd InsertDefaultMember
        token.tokPCodeSubType = INVOKE_FUNC Or INVOKE_PROPERTYGET
        output_stack.Add token
'            Dim dm As vbToken
'            Dim m As Boolean
'            m = False
        ' Is it Text1(1).Text (no default) or Text1(1)
'            if not iseol(tokens) then If tokens.Item(1).tokType = tokmember Then m = True
        Dim rtdt As TliVarType
#If 1 Then
        ' Have to insert PUT/PUTREF default member on assignments - c(1) = c(2)
'        Dim t As vbToken
        Set token = InsertDefaultMember(ii, ik, output_stack, arg_stack, mi, rtdt, rtii)
'        If token Is Nothing Then Set token = InsertDefaultMember(ii, INVOKE_FUNC Or INVOKE_PROPERTYGET, output_stack, arg_stack, mi, rtdt, rtii)
        If token Is Nothing Then
            If ik And Not (INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF) Then
                Print #99, "Required default member not found in typelib for interface: "; ii.Name
                MsgBox "Required default member not found for in typelib for interface: " & ii.Name
#If 1 Then
                Err.Raise 1 ' Default member doesn't exist"
#Else
                ' This doesn't explain OutLook's Items.Item(1).Subject --> Debug.Print Items(1)
                ' Item isn't default member for interface Items in typelib, but is used. Same with MailItem interface for Subject.
                For Each mi In ii
                    If UCase(mi.Name) = "ITEM" And (mi.InvokeKind And ik) Then Exit For
                Next
                If mi Is Nothing Then Err.Raise 1 ' default member or "Item" not found
                Print #99, "Using Item as default member name."
                MsgBox "Using Item as default member name."
#End If
            End If
            Set token = InsertDefaultMember(ii, INVOKE_FUNC Or INVOKE_PROPERTYGET, output_stack, arg_stack, mi, rtdt, rtii)
''''            Set rtii = ii
''''            ProcessArgs rtii, token, tokens, arg_stack, output_stack, INVOKE_FUNC Or INVOKE_PROPERTYGET
            If token Is Nothing Then ik = 0: Exit Sub
'            Set mi = token.tokMemberInfo
'        Else
'            Set token = t
        End If
        FunctionArgs token, tokens, arg_stack, output_stack, ik
''''        output_stack.Add t
''''        TLIFunctionArgs token, mi, arg_stack, output_stack
    Loop Until rtii Is Nothing Or ((ik And mi.InvokeKind) And arg_stack.Count = 0)
#Else
        Set token = InsertDefaultMember(ii, INVOKE_FUNC Or INVOKE_PROPERTYGET, output_stack, arg_stack, mi, rtdt, rtii)
        TLIFunctionArgs token, mi, arg_stack, output_stack
    Loop Until rtii Is Nothing Or arg_stack.Count = 0
#End If
''''    ik = mi.InvokeKind
    End If
End If
Print #99, "ProcessArgs: done"
End Sub

Sub FunctionArgs(ByRef token As vbToken, ByVal tokens As Collection, ByVal arg_stack As Collection, ByVal output_stack As Collection, ByRef ik As InvokeKinds)
Print #99, "FunctionArgs: ts="; token.tokString; " t="; token.tokType; " asc="; arg_stack.Count; " osc="; output_stack.Count; " ik="; ik
Select Case token.tokType
Case tokDeclare
''''If Not token.tokDeclare Is Nothing Then
'    ik = ik And token.tokLocalFunction.InvokeKind
    LocalFunctionArgs token, token.tokDeclare.dclParams, arg_stack, output_stack
Case tokReferenceClass, tokDeclarationInfo, tokFormClass
    Print #99, "ii="; Not token.tokInterfaceInfo Is Nothing; " mi="; Not token.tokMemberInfo Is Nothing; " pcst="; token.tokPCodeSubType
    If Not token.tokMemberInfo Is Nothing Then
        Print #99, "FunctionArgs: memid="; Hex(token.tokMemberInfo.MemberId); " off="; token.tokMemberInfo.VTableOffset; " ik="; token.tokMemberInfo.InvokeKind
        ' vbCrLf is INVOKE_UNKNOWN
        ' If token.tokMemberInfo.VTableOffset = -1 And (token.tokMemberInfo.InvokeKind <> INVOKE_FUNC And token.tokMemberInfo.InvokeKind <> INVOKE_CONST And token.tokMemberInfo.InvokeKind <> INVOKE_UNKNOWN) Then
        '    Print #99, "FunctionArgs: TLI Dispatch c="; token.tokMemberInfo.Parameters.count
        '    DispatchFunctionArgs token, tokens, arg_stack, output_stack
        ' Else
        Print #99, "FunctionArgs: TLI c="; token.tokMemberInfo.parameters.Count
        Dim rtdt As TliVarType
        Dim rtii As InterfaceInfo
'        If token.tokMemberInfo.InvokeKind <> INVOKE_UNKNOWN Then
        ik = token.tokMemberInfo.InvokeKind
        TLIFunctionArgs token, token.tokMemberInfo, arg_stack, output_stack
        If Not token.tokInterfaceInfo Is Nothing Then
            GetReturnType token.tokInterfaceInfo, token.tokMemberInfo, rtdt, rtii
            If arg_stack.Count = 0 Then
                Set token.tokVariable.varType.dtInterfaceInfo = rtii
            Else
                If rtii Is Nothing Then
                    ' Form1.Controls(1).Value = ""
                    ' UnLoad Forms(1)
                    ProcessArgStack token, tokens, arg_stack, output_stack
                Else
                    Set token.tokVariable.varType.dtInterfaceInfo = rtii
                    GoTo 20
                End If
            End If
        End If
    ElseIf Not token.tokInterfaceInfo Is Nothing Then
        ' Function references must not generate default members - RecSet(1) is not RecSet.Item(1)
        If token.tokPCodeSubType <> INVOKE_FUNC Then
            ' Dim c As Collection : c(1) (default is .Item)
            ' fixme: this isn't in the right place
            Set rtii = token.tokInterfaceInfo
20
            If Not token.tokLocalFunction Is Nothing Then Err.Raise 1
            ik = IIf(token Is LastQualifier And Not AssignmentOperator Is Nothing, INVOKE_PROPERTYPUT, INVOKE_FUNC Or INVOKE_PROPERTYGET)
            ProcessArgs rtii, token, tokens, arg_stack, output_stack, ik
        End If
    Else
        Err.Raise 1
    End If
Case tokProjectClass, tokLocalModule, tokGlobalModule
#If 1 Then
    GoTo IsVar
#Else
    Print #99, "tlf.p="; Not token.tokLocalFunction.procParams Is Nothing
    If token.tokLocalFunction.procParams Is Nothing Then Err.Raise 1
    ' fixme: perhaps variable shouldn't procParams - s/b vbext_mt_Property?
    Print #99, "tlf.m="; token.tokLocalFunction.MemberType
    ' this occurs for Form controls
    If token.tokLocalFunction.MemberType = vbext_mt_Variable Then GoTo IsVar
    If token.tokLocalFunction.MemberType <> vbext_mt_Event And token.tokLocalFunction.MemberType <> vbext_mt_Method And token.tokLocalFunction.MemberType <> vbext_mt_Property Then Err.Raise 1
    LocalFunctionArgs token, token.tokLocalFunction.procParams, arg_stack, output_stack
#End If
Case tokVariable, tokArrayVariable, tokWithValue
IsVar:
    Print #99, "FunctionArgs: v: v="; token.tokVariable.varSymbol; " mt="; token.tokVariable.MemberType; " tlf="; Not token.tokLocalFunction Is Nothing
    Print #99, "ii="; Not token.tokVariable.varType.dtInterfaceInfo Is Nothing; " class="; Not token.tokVariable.varType.dtClass Is Nothing
    If Not token.tokLocalFunction Is Nothing Then
        Print #99, "tlf="; token.tokLocalFunction.procName; " ik="; token.tokLocalFunction.InvokeKind
        ' Function references must not generate default members - RecSet(1) is not RecSet.Item(1)
''''        If token.tokLocalFunction.InvokeKind = INVOKE_FUNC Then GoTo 11
        ' new - Property Get/Let/Set shouldn't drop to Process Args where default member could be generated
        ' not sure about arg checks - ToolbarSrc\Toolbar2.vbp
        Print #99, "asc="; arg_stack.Count; " tlf.pc="; token.tokLocalFunction.procParams.Count
        If arg_stack.Count <= token.tokLocalFunction.procParams.Count Then GoTo 11
    End If
    If Not token.tokVariable.varType.dtInterfaceInfo Is Nothing Then ' TypeLib Ref var?
'        If Not token.tokMemberInfo Is Nothing Then Err.Raise 1
        ik = IIf(token Is LastQualifier And Not AssignmentOperator Is Nothing, INVOKE_PROPERTYPUT, INVOKE_FUNC Or INVOKE_PROPERTYGET)
        ProcessArgs token.tokVariable.varType.dtInterfaceInfo, token, tokens, arg_stack, output_stack, ik
    ElseIf Not token.tokVariable.varType.dtClass Is Nothing Then
        Print #99, "FunctionArgs: v: dt="; token.tokDataType; " c="; token.tokVariable.varType.dtClass.Name; " def="; token.tokVariable.varType.dtClass.DefaultMember
''''        If token.tokVariable.varType.dtClass.DefaultMember <> "" Then
' vcollection\Project1.vbp - only process default member if object
' suspect code. Should all cases of Obj with default property trigger insertion of default?
        If IsObj(token.tokDataType) And token.tokVariable.varType.dtClass.DefaultMember <> "" Then
        ' compiled class has default member
            If arg_stack.Count > 0 Then
'                If token.tokVariable.varType.dtClass.DefaultMember <> "" Then GoTo 10
                ' variable ik is 14, specify property ik for invoking property get
                ' test procparams in case of c(1) where c is a collection class with default "Item" - prjmenu132\prjmenu.vbp
                Print #99, "tlf="; Not token.tokLocalFunction Is Nothing
                If Not token.tokLocalFunction Is Nothing Then Print #99, "tlf.pp="; Not token.tokLocalFunction.procParams Is Nothing
                If Not token.tokLocalFunction Is Nothing Then If Not token.tokLocalFunction.procParams Is Nothing Then LocalFunctionArgs token, token.tokLocalFunction.procParams, arg_stack, output_stack
                ' fixme: suspicous - shouldn't tokPCodeSubType always be a property get when followed by default??? esp. if looking for Let/Set
                token.tokPCodeSubType = ik
                output_stack.Add token
                ' if getting an obj variable with a default, then only INVOKE_PROPERTYGET is set, must lookup up using both FUNC and GET
                If ik = INVOKE_PROPERTYGET Then ik = INVOKE_FUNC Or INVOKE_PROPERTYGET
                Dim p As proctable
                On Error Resume Next
                Set p = token.tokVariable.varType.dtClass.procs.Item(SymIK(token.tokVariable.varType.dtClass.DefaultMember, ik))
                ' new - SetDefault1.vbp
                If p Is Nothing Then Set p = token.tokVariable.varType.dtClass.procs.Item(SymIK(token.tokVariable.varType.dtClass.DefaultMember, INVOKE_FUNC Or INVOKE_PROPERTYGET))
                On Error GoTo 0
                '
                If p Is Nothing Then
                    Print #99, "Default member doesn't exist for """; token.tokVariable.varType.dtClass.Name; "."; token.tokVariable.varType.dtClass.DefaultMember; """ with InvokeKind of "; ik
                    MsgBox "Default member doesn't exist for """ & token.tokVariable.varType.dtClass.Name & "." & token.tokVariable.varType.dtClass.DefaultMember & """ with InvokeKind of " & ik
                    Err.Raise 1 ' default member doesn't exist for ik
                End If
                ik = p.InvokeKind
                Print #99, " def.ik="; ik; " mt="; token.tokVariable.MemberType
                Set token = InsertDefaultProjectClassMember(p)
                LocalFunctionArgs token, token.tokLocalFunction.procParams, arg_stack, output_stack
'                If arg_stack.count > 0 Then Set token = Scope.StackVariable(token, tokens, arg_stack, output_stack, ik, False, False)
                ' recursive
''''                If arg_stack.count > 0 Then FunctionArgs token, tokens, arg_stack, output_stack, ik
'                FunctionArgs token, tokens, arg_stack, output_stack, ik
            End If
        ElseIf Not token.tokLocalFunction Is Nothing Then
11
            LocalFunctionArgs token, token.tokLocalFunction.procParams, arg_stack, output_stack
        Else
            ProcessArgStack token, tokens, arg_stack, output_stack
        End If
    Else
        If Not token.tokLocalFunction Is Nothing Then
            LocalFunctionArgs token, token.tokLocalFunction.procParams, arg_stack, output_stack
        Else 'If token.tokType = tokIDispatchInterface Or (token.tokType = tokVariable And (token.tokDataType = vbObject Or token.tokDataType = VT_UNKNOWN Or (token.tokDataType = vbVariant And arg_stack.count > 0))) Then
            'If token.tokType = tokArrayVariable And Not token.tokVariable.varDimensions Is Nothing Then If token.tokVariable.varDimensions.count = 0 Then GoTo 30
            ProcessArgStack token, tokens, arg_stack, output_stack
        End If
    End If
Case tokIDispatchInterface
    If IsVObj(token.tokDataType) Then DispatchFunctionArgs token, tokens, arg_stack, output_stack
Case tokConst, tokConstantInfo, tokEnumMember, tokVariant, tokme
    ' do nothing
Case Else
    Err.Raise 1
End Select
Print #99, "FunctionArgs: done: t="; token.tokType; " dt="; token.tokDataType; " asc="; arg_stack.Count; " osc="; output_stack.Count
End Sub

Sub ProcessArgStack(ByRef token As vbToken, ByVal tokens As Collection, ByVal arg_stack As Collection, ByVal output_stack As Collection)
Print #99, "ProcessArgStack: ts="; token.tokString; " dt="; token.tokDataType; " c="; tokens.Count; " asc="; arg_stack.Count; " osc="; output_stack.Count
If arg_stack.Count > 0 Then
    Select Case token.tokDataType And Not VT_BYREF
    Case vbVariant
        ' fixme: rename DispatchFunctionArgs to ProcessVariantArgs
        ' fixme: can DispatchFunctionArgs be subsummed by ProcessArgs?
        token.tokType = tokVariantArgs
        DispatchFunctionArgs token, tokens, arg_stack, output_stack
    Case vbObject
'        token.tokType = tokInvokeDefaultMember
'        CoerceObject output_stack, output_stack.count, vbObject, Nothing, Nothing
'        output_stack(output_stack.count).tokDataType = vbVariant
        token.tokPCodeSubType = 0
        output_stack.Add token
        Set token = DispatchDefaultMember(INVOKE_FUNC Or INVOKE_PROPERTYGET)
        DispatchFunctionArgs token, tokens, arg_stack, output_stack
    Case vbUserDefinedType ' nib2000\nibnosnd.vbp - Module1.udt(1,2).mbr
        Scope.ProcessSubscripts token, arg_stack, output_stack
    Case Else
        Err.Raise 1
    End Select
End If
Print #99, "ProcessArgStack: tc="; token.tokCount; " r="; token.tokRank
End Sub

Sub LocalFunctionArgs(ByRef token As vbToken, ByVal procparms As Collection, ByVal arg_stack As Collection, ByVal output_stack As Collection)
Dim param As paramTable
Print #99, "LocalFunctionArgs: ts="; token.tokString; " pc="; procparms.Count; " asc="; arg_stack.Count; " osc="; output_stack.Count
For Each param In procparms
    Print #99, "LocalFunctionArgs:t.c="; token.tokCount; " p="; param.paramVariable.varSymbol; " at="; Hex(param.paramVariable.varAttributes); " dim="; Not param.paramVariable.varDimensions Is Nothing
    If param.paramVariable.varAttributes And VARIABLE_PUTVAL Then
        If token.tokCount + 1 <> procparms.Count Then Err.Raise 1
        ' fixme: would like to coerce putvalue on stack but Let/Set hasn't put it on yet, perhaps should
' fixme: isn't token.tokVariable same as last param varVariable? So following does nothing?
'If Not token.tokVariable Is param.paramVariable Then MsgBox "LocalFunctionArgs: internal error: PUTVAL variables differ"
' obsolete?
token.tokVariable.varAttributes = token.tokVariable.varAttributes Or (param.paramVariable.varAttributes And (VARIABLE_PUTVAL Or VARIABLE_ARRAY Or VARIABLE_BYREF))
        ' fixme: shouldn't VT_BYREF already be set? CHECK THIS OUT! property1.vbp
'        token.tokDataType = param.paramVariable.varType.dtDataType Or (param.paramVariable.varAttributes And (VARIABLE_ARRAY Or VARIABLE_BYREF))
#If 0 Then
        Print #99, "putval.dt="; Not param.paramVariable.varType.dtDataType
        Print #99, "putval.ii="; Not param.paramVariable.varType.dtInterfaceInfo Is Nothing
        Print #99, "putval.class="; Not param.paramVariable.varType.dtClass Is Nothing
        Dim dt As TliVarType
        dt = param.paramVariable.varType.dtDataType Or (param.paramVariable.varAttributes And (VARIABLE_ARRAY Or VARIABLE_BYREF))
        CoerceOperand gOptimizeFlag, output_stack, output_stack.Count, dt, param.paramVariable.varType.dtInterfaceInfo, param.paramVariable.varType.dtClass
#End If
        Exit For
    End If
    Do
        token.tokCount = token.tokCount + 1
        Print #99, " tc="; token.tokCount; " asc="; arg_stack.Count
        Dim arg As Collection
        Set arg = Nothing
        On Error Resume Next
        Set arg = arg_stack.Item(CStr(token.tokCount))
        On Error GoTo 0
        Print #99, "positional arg="; Not arg Is Nothing
        If arg Is Nothing Then
            On Error Resume Next
            Set arg = arg_stack.Item(UCase(param.paramVariable.varSymbol))
            On Error GoTo 0
            Print #99, "named arg="; Not arg Is Nothing
            If arg Is Nothing Then
                Set arg = New Collection
            Else
                arg_stack.Remove UCase(param.paramVariable.varSymbol)
            End If
        Else
            If Not arg Is arg_stack.Item(1) Then Err.Raise 1
            arg_stack.Remove 1
        End If
        Print #99, "ac="; arg.Count
        Dim dt As TliVarType
'        dt = param.paramVariable.varType.dtDataType Or (param.paramVariable.varAttributes And VT_BYREF)
        dt = param.paramVariable.varType.dtDataType Or (param.paramVariable.varAttributes And (VARIABLE_ARRAY Or VARIABLE_BYREF))
' fixme: not sure how PARAMARRAY parameter is getting varDimensions allocated. Does it matter?
        If Not CBool(param.paramVariable.varAttributes And VARIABLE_PARAMARRAY) Then If Not param.paramVariable.varDimensions Is Nothing Then dt = dt Or VT_ARRAY Or VT_BYREF
        Dim ii As InterfaceInfo
        Set ii = Nothing
        Dim m As vbModule
        Set m = Nothing
        Print #99, "dt="; dt; " type="; param.paramVariable.varType.dtType
        If (dt And Not VT_BYREF) = vbObject Then
            Select Case param.paramVariable.varType.dtType
            Case tokProjectClass
                Set m = param.paramVariable.varType.dtClass
            Case tokReferenceClass
                Set ii = param.paramVariable.varType.dtInterfaceInfo
            Case tokIDispatchInterface
                Set ii = DispatchInterfaceInfo
            Case tokFormClass
                Set ii = param.paramVariable.varType.dtInterfaceInfo
            Case Else
                Print #99, "LocalFunctionArgs: Invalid Class type: "; param.paramVariable.varType.dtType
                Err.Raise 1 ' Invalid Class type
            End Select
        End If
        Print #99, "ac="; arg.Count; " vt="; varType(param.paramVariable.varVariant); " dt="; dt; " isEmpty="; IsEmpty(param.paramVariable.varVariant)
        If arg.Count = 0 Then
#If 0 Then
            If Not param.paramVariable.varAttributes And VARIABLE_OPTIONAL Then Err.Raise 1 ' internal error - expecting param
            output_stack.Add MissingArg(New vbToken, vbVariant Or param.paramVariable.varAttributes And VARIABLE_ARRAY)
#Else
            If param.paramVariable.varAttributes And VARIABLE_OPTIONAL Then
                If param.paramVariable.varAttributes And VARIABLE_DEFAULTVALUE Then
                    output_stack.Add New vbToken
                    If (dt And Not VT_BYREF) = vbObject Then
                        If Not IsEmpty(param.paramVariable.varVariant) Then Err.Raise 1 ' internal error
                        output_stack.Item(output_stack.Count).tokType = tokNothing
                    Else
                        If IsEmpty(param.paramVariable.varVariant) Then Err.Raise 1 ' internal error
                        If (dt And Not VT_BYREF) <> varType(param.paramVariable.varVariant) Then Err.Raise 1
                        output_stack.Item(output_stack.Count).tokType = tokVariant
                        output_stack.Item(output_stack.Count).tokValue = param.paramVariable.varVariant
                    End If
                    output_stack.Item(output_stack.Count).tokString = "_Param" & CStr(token.tokCount)
                    output_stack.Item(output_stack.Count).tokDataType = dt
                Else
                    If Not IsEmpty(param.paramVariable.varVariant) Then Err.Raise 1 ' internal error
                    output_stack.Add MissingArg(New vbToken, vbVariant Or (dt And VARIABLE_ARRAY))
                End If
            Else
                Err.Raise 1 ' internal error - expecting param
            End If
#End If
        Else
            Dim t As vbToken
            For Each t In arg
                Print #99, "t="; t.tokString
                output_stack.Add t
                arg.Remove 1 ' removing within For Each should be OK
            Next
        End If
'        CoerceOperand goptimizeflag, output_stack, output_stack.count, dt Or param.paramVariable.varAttributes And VARIABLE_ARRAY, ii, m
        CoerceOperand gOptimizeFlag, output_stack, output_stack.Count, dt, ii, m
    Loop While CBool(param.paramVariable.varAttributes And VARIABLE_PARAMARRAY) And arg_stack.Count > 0   ' loop for paramarray parameters
Next
Print #99, "LocalFunctionArgs: done: acs="; arg_stack.Count
'If arg_stack.count > 0 Then Err.Raise 1
End Sub

Sub TLIFunctionArgs(ByVal token As vbToken, ByVal mi As MemberInfo, ByVal arg_stack As Collection, ByVal output_stack As Collection)
Dim parameters As parameters
Print #99, "TLIFunctionArgs: t="; token.tokType; " s="; token.tokString; " dt="; token.tokDataType; " ii="; Not token.tokInterfaceInfo Is Nothing; " osc="; output_stack.Count; " asc="; arg_stack.Count; " mi.pc="; mi.parameters.Count; " mi.dk="; mi.DescKind; " mi.am="; mi.AttributeMask
'If IsVObj(token.tokDataType) Then
#If 1 Then
If token.tokType = tokReferenceClass Then ' If a member, do a QI
    If token.tokInterfaceInfo Is Nothing Then Err.Raise 1
    Print #99, "ii.am="; Hex(token.tokInterfaceInfo.AttributeMask)
    ' fixme: probably should set VARIABLE_APPOBJECT flag and test here
    If Not token.tokVariable.varType.dtClassInfo Is Nothing Then If token.tokVariable.varType.dtClassInfo.AttributeMask And TYPEFLAG_FAPPOBJECT Then GoTo 12
    If output_stack.Count = 0 Then Err.Raise 1
    ' Object.someQI(SubRef(ByRef))
    If output_stack(output_stack.Count).tokDataType And VT_BYREF Then OutputSubRef output_stack, output_stack.Count
    CoerceObject output_stack, output_stack.Count, output_stack(output_stack.Count).tokDataType, token.tokInterfaceInfo, Nothing
12
End If
#Else
Dim b As Boolean ' testing - only tokReferenceClass needs a QI -- right?
If Not token.tokInterfaceInfo Is Nothing And output_stack.Count > 0 Then If IsObj(output_stack(output_stack.Count).tokDataType) Then CoerceObject output_stack, output_stack.Count, output_stack(output_stack.Count).tokDataType, token.tokInterfaceInfo, Nothing: b = True
If b Xor (token.tokType = tokReferenceClass) Then
    MsgBox "TLIFunctionArgs: type=" & token.tokType & " b=" & b
End If
#End If
If output_stack.Count > 0 Then If Not output_stack(output_stack.Count).tokInterfaceInfo Is Nothing Then Print #99, "os(1).ii="; output_stack(output_stack.Count).tokInterfaceInfo.Name; " tk="; output_stack(output_stack.Count).tokInterfaceInfo.TypeKind
Set parameters = mi.parameters
'If output_stack.count > 0 Then If IsObj(output_stack(output_stack.count).tokDataType) Then CoerceObject output_stack, output_stack.count, token.tokInterfaceInfo, Nothing
Dim pi As ParameterInfo
For Each pi In parameters
    Print #99, "TLIFunctionArgs: c="; token.tokCount; " pi="; pi.Name; " vt="; pi.VarTypeInfo.varType; " pf="; Hex(pi.flags); " pl="; pi.VarTypeInfo.PointerLevel
    Dim dt As VariantTypeConstants
    Dim vt As vbDataType
    Set vt = New vbDataType
    dt = InitVarTypeFromVarTypeInfo(pi.VarTypeInfo, vt)
    Print #99, "dt="; dt
    ' MSDASC.DataLinks.PromptEdit has a FOUT parameter but VT_BYREF isn't set, so needs setting
    If pi.flags And PARAMFLAG_FOUT Then
        If dt And Not VT_BYREF Then
            Print #99, "ByRef parameter missing ByRef flag:"; mi.Name; "."; pi.Name
'            MsgBox "ByRef parameter missing ByRef flag:" & mi.Name & "." & pi.Name
            dt = dt Or VT_BYREF
        End If
    End If
    If pi.Name = "" Then Print #99, "TLIFunctionArgs: ParameterInfo.Name is blank!"
    If pi.flags And PARAMFLAG_FLCID Then GoTo Continue
    If pi.Optional Xor CBool(pi.flags And PARAMFLAG_FOPT) Then
        ' didn't work on move 1,2
        Print #99, "warning: TLIFunctionArgs: Optional <> (Flags and FOPT): f=" & Hex(pi.flags)
'        MsgBox "warning: TLIFunctionArgs: Optional <> (Flags and FOPT): f=" & Hex(ParameterInfo.Flags)
    End If
    Print #99, "TLIFunctionArgs: 5 f="; pi.flags And PARAMFLAG_FRETVAL
    ' removed PARAMFLAG_FOUT - optional parameter used it
    If pi.flags And PARAMFLAG_FRETVAL Then Exit For
    Print #99, "TLIFunctionArgs: 6"
    Do
' need to exit for if nothing to be processed - needs work
' kludge - exit on lcid idl flag - assumes nothing of interest follows (should LCID be skipped instead?)
'        If arg_stack.Item(1).tokString = ")" Then Exit For
' Form1.Controls has FRETVAL parameter, can't seem to get VTable? description
        Print #99, "TLIFunctionArgs: 7 tc: "; token.tokCount; " acs="; arg_stack.Count; " ik="; mi.InvokeKind
        If Not mi.ReturnType Is Nothing Then Print #99, "rt="; mi.ReturnType.varType
        If token.tokCount = parameters.Count And (mi.InvokeKind And (INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF)) And (mi.ReturnType.varType = VT_VOID Or mi.ReturnType.varType = VT_HRESULT) Then Exit Do
        token.tokCount = token.tokCount + 1 ' don't count return type
        Dim arg As Collection
        Set arg = Nothing
        On Error Resume Next
        Set arg = arg_stack.Item(CStr(token.tokCount))
        On Error GoTo 0
        Print #99, "TLIFunctionArgs: 8 positional arg="; Not arg Is Nothing
        If arg Is Nothing Then
            On Error Resume Next
            Set arg = arg_stack.Item(UCase(pi.Name))
            On Error GoTo 0
            Print #99, "TLIFunctionArgs: 9 named arg="; Not arg Is Nothing
            ' this can happen for propput/propputrefs - putval is already on stack
            If arg Is Nothing Then
                If token.tokCount <= arg_stack.Count Then Err.Raise 1
                token.tokCount = token.tokCount - 1
                Exit For
            End If
            arg_stack.Remove UCase(pi.Name)
#If 0 Then
        Else
            If Not arg Is arg_stack.Item(1) Then Err.Raise 1
            arg_stack.Remove 1
#End If
        End If
        Print #99, "TLIFunctionArgs: 10 tc="; token.tokCount; " ac="; arg.Count; " osc="; output_stack.Count; " pst="; token.tokPCodeSubType
        If arg.Count = 0 Then Err.Raise 1
        Dim t As vbToken
        For Each t In arg
            Print #99, "t="; t.tokString
            output_stack.Add t
            arg.Remove 1 ' removing within For Each should be OK
        Next
        If IsParamArrayArg(mi, token.tokCount) Then If dt <> (VT_VARIANT Or VT_ARRAY Or VT_BYREF) Then Err.Raise 1
        CoerceOperand gOptimizeFlag, output_stack, output_stack.Count, IIf(IsParamArrayArg(mi, token.tokCount), VT_VARIANT, dt), vt.dtInterfaceInfo
Dim pl As Integer
pl = pi.VarTypeInfo.PointerLevel
If pl < 0 Then MsgBox "TLIFunctionArgs: pl=" & pl
If IsObj(dt) Then pl = pl - 1
If dt And VT_BYREF Then pl = pl - 1
Print #99, "TLIFunctionArgs: ts=" & token.tokString & " n=" & pi.Name & " pl=" & pl
Dim i As Integer
For i = 1 To pl
OutputAddRef output_stack, output_stack.Count
Next
        Dim rm As Boolean
        rm = True
    Loop While IsParamArrayArg(mi, token.tokCount) And token.tokCount < arg_stack.Count  ' loop for paramarray parameters
Continue:
Next
Print #99, "TLIFunctionArgs: done"
#If 0 Then
If arg_stack.Count > 0 Then Err.rase 1
#Else
' If any args are processed, then all args should be processed?
If rm Then
    Dim tokens As Collection
    For Each tokens In arg_stack
        ' fixme: Is is better to do RemoveAllTokens to clear nested collection?
        'RemoveAllTokens tokens
        arg_stack.Remove 1
    Next
End If
#End If
End Sub

Sub DispatchFunctionArgs(ByRef token As vbToken, ByVal tokens As Collection, ByVal arg_stack As Collection, ByVal output_stack As Collection)
Print #99, "DispatchFunctionArgs: ts="; token.tokString; " dt="; token.tokDataType; " tc="; token.tokCount; " ii="; Not token.tokInterfaceInfo Is Nothing; " asc="; arg_stack.Count; " osc="; output_stack.Count
' Hmmm, should tokType be changed (from tokVariable) or should sub-type be used?
'token.tokType = tokIDispatchInterface
#If 0 Then ' not here, coerce when generating default member
If IsObj(token.tokDataType) Then
' temporarily put token on stack, coerce to object, coerce to IDispatch interface
' Unload Form1 - has output_stack.count = 0
    output_stack.Add token
'' Should object expression be coerced to object?
'    CoerceOperand goptimizeflag, output_stack, output_stack.count, vbObject
    CoerceObject gOptimizeFlag, output_stack, output_stack.Count, vbObject, Nothing, Nothing
    Set token = output_stack.Item(output_stack.Count)
    output_stack.Remove output_stack.Count
End If
#End If
'If token.tokDataType <> vbVariant Then Err.Raise 1
While arg_stack.Count > 0
    token.tokCount = token.tokCount + arg_stack.Count
    Dim arg As Collection
    ' fixme: implemented named arguments for Invoke
    Dim i As Integer
    For i = 1 To arg_stack.Count
        Set arg = Nothing
        On Error Resume Next
        Set arg = arg_stack.Item(i)
        On Error GoTo 0
        If arg Is Nothing Then
            MsgBox "DispatchFunctionArgs: named arguments for dispatch objects not implemented"
            Err.Raise 1
        End If
    Next
    For Each arg In arg_stack
        If arg.Count = 0 Then
            output_stack.Add MissingArg(New vbToken, vbVariant)
        Else
            Dim t As vbToken
            For Each t In arg
                output_stack.Add t
            Next
            CoerceOperand gOptimizeFlag, output_stack, output_stack.Count, VT_VARIANT
        End If
        arg_stack.Remove 1 ' should be ok to remove within For Each
    Next
    ' Look for more arguments, but not if token is a special arg processing type
    ' such as (Print/Line/...) which won't have "(" following. change3g\definear.vbp
    If Not IsEOL(tokens) Then If tokens.Item(1).tokString = "(" Then Set arg_stack = Scope.CreateArgStack(token, tokens, New Collection)
Wend
Print #99, "DispatchFunctionArgs: done: osc="; output_stack.Count
'output_stack.Add token
'Set token = MakeInvoke
End Sub

Function MakeInvoke() As vbToken
Set MakeInvoke = New vbToken
MakeInvoke.tokType = tokInvoke
MakeInvoke.tokString = "Invoke"
MakeInvoke.tokDataType = vbVariant
End Function

' maybe should use VT_ERROR
Function MissingArg(ByVal token As vbToken, ByVal vt As TliVarType) As vbToken
Print #99, "MissingArg: ts="; token.tokString; " vt="; vt
token.tokString = "_Missing"
token.tokType = tokMissing
token.tokDataType = vt
Set MissingArg = token
End Function

' fixme: eliminate tokAddRef, tokSubRef by implementing PointerLevel
Sub OutputAddRef(ByVal output_stack As Collection, ByRef cnt As Long)
Print #99, "OutputAddRef: s="; output_stack(cnt).tokString, " osc="; output_stack.Count; " cnt="; cnt
Dim token As New vbToken
token.tokString = "AddRef"
token.tokType = tokAddRef
Set token.tokInterfaceInfo = output_stack.Item(cnt).tokInterfaceInfo
Set token.tokMemberInfo = output_stack.Item(cnt).tokMemberInfo
Set token.tokVariable = output_stack.Item(cnt).tokVariable
token.tokDataType = output_stack.Item(cnt).tokDataType Or VT_BYREF
output_stack.Add token, , , cnt
cnt = cnt + 1
End Sub

Sub OutputSubRef(ByVal output_stack As Collection, ByRef cnt As Long)
Print #99, "OutputSubRef: s="; output_stack(cnt).tokString, " osc="; output_stack.Count; " cnt="; cnt
Dim token As New vbToken
token.tokString = "SubRef"
token.tokType = tokSubRef
Set token.tokInterfaceInfo = output_stack.Item(cnt).tokInterfaceInfo
Set token.tokMemberInfo = output_stack.Item(cnt).tokMemberInfo
Set token.tokVariable = output_stack.Item(cnt).tokVariable
token.tokDataType = output_stack.Item(cnt).tokDataType And Not VT_BYREF
output_stack.Add token, , , cnt
cnt = cnt + 1
End Sub

Sub getFileNumber(ByVal tokens As Collection, output_stack As Collection)
Dim special As SpecialTypes
special = getSpecialTypes(tokens.Item(1))
If special = SPECIAL_NS Then tokens.Remove 1
oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbInteger
End Sub

Sub AddLabelRef(ByVal token As vbToken, output_stack As Collection)
On Error Resume Next
Set token.tokLabel = currentProc.procLabels.Item(UCase(token.tokString))
On Error GoTo 0
If token.tokLabel Is Nothing Then
    currentProc.procLabels.Add New vbLabel, UCase(token.tokString)
    Set token.tokLabel = currentProc.procLabels.Item(currentProc.procLabels.Count)
    token.tokLabel.labSymbol = token.tokString
End If
token.tokLabel.labRefs.Add token
token.tokType = tokLabelRef
output_stack.Add token
End Sub

Sub AddLabelDef(ByVal token As vbToken, output_stack As Collection)
On Error Resume Next
Set token.tokLabel = currentProc.procLabels.Item(UCase(token.tokString))
On Error GoTo 0
If token.tokLabel Is Nothing Then
    currentProc.procLabels.Add New vbLabel, UCase(token.tokString)
    Set token.tokLabel = currentProc.procLabels.Item(currentProc.procLabels.Count)
    token.tokLabel.labSymbol = token.tokString
Else
    If token.tokLabel.labAddr <> 0 Then Err.Raise 1 ' duplicate label definition
End If
token.tokLabel.labAddr = 123 ' What label address should go here?
token.tokType = tokLabelDef
output_stack.Add token
currentProc.procStatements.Add output_stack
End Sub

Sub EndProc()
Print #99, "EndProc: cp="; Not currentProc Is Nothing
Dim l As vbLabel
For Each l In currentProc.procLabels
    If l.labAddr = 0 Then MsgBox "Label " & l.labSymbol & " not defined": Err.Raise 1
Next
Dim n As Object
For Each n In currentProc.procNests
    MsgBox n.stmtTok.tokString & " not terminated"
Next
Set currentProc = Nothing
End Sub

Function ProcAdd(ByVal procName As String, ByVal attr As procattributes, ByVal MemberType As vbext_MemberType, ByVal InvokeKind As InvokeKinds) As proctable
Dim ProcNameIK As String
Print #99, "ProcAdd: n="; procName; " a="; Hex(attr); " currentmodule="; currentModule.Name; " membertype="; MemberType; " ik="; InvokeKind
If UCase(procName) = "ERROR" Then
    ' Possible bug or misfeature in VB allowing Error to be used as a method or variable identifier.
    Print #99, """Error"" is a statement identifier and may not be used as a method or variable identifier. Please rename."
    MsgBox """Error"" is a statement identifier and may not be used as a method or variable identifier. Please rename."
    Err.Raise 1
End If
Set ProcAdd = New proctable
Print #99, "1"
ProcAdd.procName = procName
Print #99, "2"
ProcAdd.procattributes = attr Or proc_attr_defined
Print #99, "3"
Set ProcAdd.procLocalModule = currentModule
If currentModule Is Nothing Then Err.Raise 1
Print #99, "4"
ProcAdd.MemberType = MemberType
ProcAdd.InvokeKind = InvokeKind
ProcNameIK = SymIK(procName, IIf(InvokeKind = INVOKE_FUNC Or InvokeKind = INVOKE_PROPERTYGET, INVOKE_FUNC Or INVOKE_PROPERTYGET, InvokeKind))
Print #99, "5 pnik="; ProcNameIK
Select Case MemberType
Case vbext_mt_Method, vbext_mt_Variable, vbext_mt_Property
    If currentModule.procs.Count = 0 Then currentModule.procs.Add ProcAdd, ProcNameIK Else currentModule.procs.Add ProcAdd, ProcNameIK, 1
Case vbext_mt_Event
    If currentModule.events.Count = 0 Then currentModule.events.Add ProcAdd, ProcNameIK Else currentModule.events.Add ProcAdd, ProcNameIK, 1
Case Else
    Print #99, "ProcAdd: Unknown MemberType="; MemberType
    Err.Raise 1
End Select
Print #99, "6"
Dim pt As proctable
For Each pt In currentModule.procs
  Print #99, "ProcAdd: ppt="; pt.procName;
  If Not pt.procFunctionResultType Is Nothing Then Print #99, " dt="; pt.procFunctionResultType.varType.dtDataType;
  Print #99,
Next
For Each pt In currentModule.events
  Print #99, "ProcAdd: ept="; pt.procName
Next
Print #99, "7"
End Function

Function checkNesting(ByVal n As Collection) As Boolean
If n.Count = 0 Then checkNesting = True: Exit Function
If currentProc.procNests.Count = 0 Then Err.Raise 1 ' internal error
If Not n.Item(1) Is currentProc.procNests.Item(1) Then checkNesting = True: Exit Function ' end without start or invalid nesting
End Function

Sub AppendCollection(ByVal colIn As Collection, ByVal colOut As Collection)
Dim v As Variant
For Each v In colIn
    colOut.Add v
'    colIn.remove 1
Next
End Sub

' fixme: StmtMethod in RPNize may not be correct for some calls - paste in RPNize instead? (like StmtCall)
' fixme: obsolete this routine?
Function SymbolLookUp(ByVal OptimizeFlag As Long, ByVal tokens As Collection, ByVal output_stack As Collection, ByVal ik As InvokeKinds, Optional ByVal StmtMethod As Boolean, Optional ByVal LetStmt As Boolean, Optional ByVal NoInsertObjDefault As Boolean) As vbToken
Dim token As vbToken
Dim Scope As vbtScope

If IsEOL(tokens) Then
    Print #99, "Expecting symbol"
    MsgBox "Expecting symbol"
    Err.Raise 1
End If

Print #99, "SymbolLookUp: opt="; OptimizeFlag; " c="; tokens.Count; " os="; output_stack.Count; " ik="; ik; " sm="; StmtMethod
For Each token In tokens
    Print #99, "SymbolLookUp: before: tok="; token.tokString; " t="; token.tokType
Next
For Each token In output_stack
    Print #99, "SymbolLookUp: before: os: s="; token.tokString; " t="; token.tokType; " dt="; token.tokDataType; " st="; token.tokPCodeSubType
Next

oRPN.RPNize OptimizeFlag, tokens, output_stack, -1, ik, StmtMethod, LetStmt, NoInsertObjDefault
' fix return so this doesn't have to be done

For Each token In output_stack
    Print #99, "SymbolLookUp: after: os: s="; token.tokString; " t="; token.tokType; " dt="; token.tokDataType; " st="; token.tokPCodeSubType
Next

' probably should check that SymbolLookUp has been assigned a symbol - tokmember, ...
Set SymbolLookUp = output_stack.Item(output_stack.Count)
output_stack.Remove output_stack.Count

For Each token In tokens
Print #99, "SymbolLookUp: after: tok="; token.tokString
Next
End Function

' DOES NOT ADD VARIABLE TO VARIABLE TABLE!
Function CreateVariable(ByVal s As String, ByVal attr As VariableAttributes) As vbVariable
Print #99, "CreateVariable: s="; s; " attr="; attr
Set CreateVariable = New vbVariable
CreateVariable.MemberType = vbext_mt_Variable
Set CreateVariable.varType = New vbDataType
CreateVariable.varSymbol = s
CreateVariable.varAttributes = attr
Set CreateVariable.varModule = currentModule
Set CreateVariable.varProc = currentProc
End Function

' Note: VB has no method for implicitly declaring a variable as an array
Function CreateVariant(ByVal tokens As Collection, Optional ByVal ReDimFlag As Boolean) As vbToken
Print #99, "CreateVariant: t.c="; tokens.Count; " rdf="; ReDimFlag; " pn="; PassNumber
Set CreateVariant = tokens.Item(1)
' temporarily commented out - until member and argument processing is done
#If 1 Then
If PassNumber = 1 Then
' reference must be a not-yet-defined Const or Enum member.
' Just create a empty variant for reference. Later pass will evaluate.
    CreateVariant.tokType = tokVariant
Else
    If currentModule.OptionExplicit And Not ReDimFlag Then
        Print #99, "Variable not explicitly defined: " & CreateVariant.tokString
        MsgBox "Variable not explicitly defined: " & CreateVariant.tokString
        Err.Raise 1
    End If
#End If
    Print #99, "CreateVariant: creating "; CreateVariant.tokString
    CreateVariant.tokType = tokVariable
    Print #99, "1"
    Set CreateVariant.tokModule = currentModule
    Print #99, "2"
    Set CreateVariant.tokVariable = New vbVariable
    Print #99, "3"
    CreateVariant.tokVariable.MemberType = vbext_mt_Variable
    Print #99, "4"
    currentProc.procLocalVariables.Add CreateVariant.tokVariable, UCase(CreateVariant.tokString)
    Print #99, "5"
    ' tokDataType <> 0 for undeclared variables with explicit type specifier (e.g. A$, I%)
    If CreateVariant.tokDataType = 0 Then CreateVariant.tokDataType = GetDefaultType(CreateVariant.tokString)
    If ReDimFlag Then
        Set CreateVariant.tokVariable.varType = Nothing
        Set CreateVariant.tokVariable.varDimensions = New Collection ' ReDim variable must have array subscripts
    Else
        Print #99, "8 dt="; CreateVariant.tokDataType
        Set CreateVariant.tokVariable.varType = New vbDataType
        Print #99, "9"
        CreateVariant.tokVariable.varType.dtDataType = CreateVariant.tokDataType
        Print #99, "10"
    End If
    CreateVariant.tokVariable.varSymbol = CreateVariant.tokString
    Print #99, "11"
    Set CreateVariant.tokVariable.varModule = currentModule
    Print #99, "12"
    Set CreateVariant.tokVariable.varProc = currentProc
End If
Print #99, "13"
tokens.Remove 1
Print #99, "14"
End Function

Function GetDefaultType(ByVal c As String) As TliVarType
GetDefaultType = vbVariant
On Error Resume Next
GetDefaultType = currentModule.DefaultTypes.Item(UCase(Left(c, 1)))
End Function

Function HasValue(ByVal token As vbToken) As Boolean
#If 0 Then
HasValue = IsAny(token.tokType, tokVariant, tokNothing)
If Not HasValue And PassNumber > 1 Then HasValue = IsAny(token.tokType, tokConst, tokConstantInfo, tokEnumMember)
#Else
HasValue = IsAny(token.tokType, tokVariant, tokNothing, tokConst, tokConstantInfo, tokEnumMember)
#End If
Print #99, "HasValue: hv="; HasValue; " e="; IsEmpty(token.tokValue)
If PassNumber = 1 Then Err.Raise 1 ' Internal error
' an undefined precompiled constant symbol HasValue and IsEmpty
'If HasValue Eqv IsEmpty(token.tokValue) Then Err.Raise 1
End Function

Sub RemoveAllTokens(ByVal tokens As Collection)
Dim token As vbToken
For Each token In tokens
    If token.tokString = ":" Then Exit For
    tokens.Remove 1
Next
End Sub

Sub SkipToNextComma(ByVal tokens As Collection)
Print #99, "SkipToNextComma: t.c="; tokens.Count
Dim token As vbToken
For Each token In tokens
    If token.tokString = "," Or token.tokString = ":" Then Exit For
    tokens.Remove 1
    If token.tokString = ")" Then Exit For
    If token.tokString = "(" Then SkipToNextComma tokens
Next
Print #99, "SkipToNextComma: done: t.c="; tokens.Count
End Sub

Function IsDefaultMember(ByVal mi As MemberInfo, ByVal ik As InvokeKinds) As Boolean
' Can't rely on (.MemberId mod 65536) if implied interfaces are used, might erroneously match
'IsDefaultMember = mi.MemberId = 0 And CBool(mi.InvokeKind And ik)
' trying Mod again - only use for Dispatch interfaces
' trying (mi.AttributeMask and 1) to reject IDispatch and IUnknown interfaces - don't want their members to be defaults
Print #99, "IsDefaultMember: am="; Hex(mi.AttributeMask)
IsDefaultMember = (mi.AttributeMask And 1) = 0 And (mi.MemberId Mod 65536) = 0 And CBool(mi.InvokeKind And ik)
End Function

Function InitVarTypeFromVarTypeInfo(ByVal vti As VarTypeInfo, ByVal vt As vbDataType) As TliVarType
Print #99, "InitVarTypeFromVarTypeInfo: vti exists="; Not vti Is Nothing
If vti Is Nothing Then Err.Raise 1
On Error Resume Next
Print #99, "vti.vt="; vti.varType
Print #99, "vti.pl="; vti.PointerLevel
Print #99, "vti.IsExt="; vti.IsExternalType
Print #99, "vti.ti.tk="; vti.TypeInfo.TypeKind
Print #99, "vti.ti.n="; vti.TypeInfo.Name
Print #99, "vti.ti.am="; Hex(vti.TypeInfo.AttributeMask)
Print #99, "vti.ti.g="; vti.TypeInfo.GUID
Print #99, "vti.rt.vt"; vti.TypeInfo.ResolvedType.varType
Print #99, "vti.eTLib.n="; vti.TypeLibInfoExternal.Name
Print #99, "vti.eTLib.ti.n="; vti.TypeLibInfoExternal.TypeInfos(vti.TypeInfoNumber).Name
Print #99, "vti.eTLib.ti.am="; Hex(vti.TypeLibInfoExternal.TypeInfos(vti.TypeInfoNumber).AttributeMask)
Print #99, "vti.eTLib.ti.g="; vti.TypeLibInfoExternal.TypeInfos(vti.TypeInfoNumber).GUID
Print #99, "vti.eTLib.ti.tk="; vti.TypeLibInfoExternal.TypeInfos(vti.TypeInfoNumber).TypeKind
Print #99, "vti.eTLib.ti.rt.vt="; vti.TypeLibInfoExternal.TypeInfos(vti.TypeInfoNumber).ResolvedType.varType
On Error GoTo 0
'If vti.PointerLevel > 1 Then Err.Raise 1
InitVarTypeFromVarTypeInfo = vti.varType
If (InitVarTypeFromVarTypeInfo And Not (VT_ARRAY Or VT_BYREF Or VT_VECTOR)) = 0 Then
Print #99, "8 tk="; vti.TypeInfo.TypeKind
' using "dt Or" atleast needed for VT_VECTOR (possible with TKIND_RECORD), VT_ARRAY or VT_BYREF
        Select Case vti.TypeInfo.TypeKind
            Case TKIND_ENUM
                Set vt.dtConstantInfo = vti.TypeInfo
                vt.dtType = tokReferenceClass
                Print #99, "ci="; vt.dtConstantInfo.Name
                InitVarTypeFromVarTypeInfo = InitVarTypeFromVarTypeInfo Or vbLong
            Case TKIND_COCLASS
                Set vt.dtInterfaceInfo = vti.TypeInfo.DefaultInterface
                vt.dtType = tokReferenceClass
                Print #99, "ii="; vt.dtInterfaceInfo.Name
                InitVarTypeFromVarTypeInfo = InitVarTypeFromVarTypeInfo Or vbObject
            Case TKIND_INTERFACE
                Set vt.dtInterfaceInfo = vti.TypeInfo
                vt.dtType = tokReferenceClass
                Print #99, "ii="; vt.dtInterfaceInfo.Name
                InitVarTypeFromVarTypeInfo = InitVarTypeFromVarTypeInfo Or VT_UNKNOWN
            Case TKIND_DISPATCH
                Set vt.dtInterfaceInfo = vti.TypeInfo
                vt.dtType = tokIDispatchInterface
                Print #99, "ii="; vt.dtInterfaceInfo.Name
                InitVarTypeFromVarTypeInfo = InitVarTypeFromVarTypeInfo Or vbObject
            Case TKIND_RECORD
                Set vt.dtRecordInfo = vti.TypeInfo
                vt.dtType = tokReferenceClass
                Print #99, "ri="; vt.dtRecordInfo.Name
                InitVarTypeFromVarTypeInfo = InitVarTypeFromVarTypeInfo Or vbUserDefinedType
            Case Else
                Print #99, "Unsupported TypeKind "; vti.TypeInfo.TypeKind
                MsgBox "Unsupported TypeKind " & vti.TypeInfo.TypeKind
                Err.Raise 1
        End Select
    End If
' fixme: remove when pl is implemented
If Not IsObj(InitVarTypeFromVarTypeInfo) And vti.PointerLevel = 1 Then InitVarTypeFromVarTypeInfo = InitVarTypeFromVarTypeInfo Or VT_BYREF
Print #99, "InitVarTypeFromVarTypeInfo: dt="; InitVarTypeFromVarTypeInfo
End Function

Function GetForm(ByVal m As vbModule) As vbToken
Dim tlib As TypeLibInfo
Dim classinfo As CoClassInfo
Print #99, "GetForm: vb path="; tliVB.FullPath
Set tlib = TypeLibInfoFromFile(tliVB.FullPath)
If tlib Is Nothing Then Err.Raise 1 ' Typelibrary error
Print #99, "GetForm: n="; m.Name; " ct="; m.Component.Type
Select Case m.Component.Type
    Case vbext_ct_VBForm
        Set classinfo = tlib.CoClasses.NamedItem("Form")
    Case vbext_ct_VBMDIForm
        Set classinfo = tlib.CoClasses.NamedItem("MDIForm")
    Case vbext_ct_PropPage
        Set classinfo = tlib.CoClasses.NamedItem("PropertyPage") ' axmrquee\axmrquee.vbp
    Case vbext_ct_UserControl
        Set classinfo = tlib.CoClasses.NamedItem("UserControl") ' axmrquee\axmrquee.vbp
    Case Else
        For Each classinfo In tlib.CoClasses
            Print #99, " ci="; classinfo.Name; " guid="; classinfo.GUID
        Next
        Err.Raise 1 ' unexpected component type
End Select
' Form1.method
' Put this coclass form code into module variable m.coclassinfo
Print #99, "4"
If classinfo Is Nothing Then Err.Raise 1
Print #99, "5"
Set tlib = Nothing
Print #99, "6"
Set GetForm = New vbToken
Print #99, "7"
If Not GetForm.tokVariable Is Nothing Then Err.Raise 1
Set GetForm.tokVariable = New vbVariable
GetForm.tokVariable.MemberType = vbext_mt_Variable
Print #99, "8"
If Not GetForm.tokVariable.varType Is Nothing Then Err.Raise 1
Set GetForm.tokVariable.varType = New vbDataType
GetForm.tokVariable.varType.dtType = tokFormClass
Print #99, "9"
Set GetForm.tokVariable.varType.dtClass = m
Set GetForm.tokVariable.varType.dtClassInfo = classinfo
Print #99, "10 ii.n="; classinfo.DefaultInterface.Name; " am="; Hex(classinfo.DefaultInterface.AttributeMask)
Set GetForm.tokVariable.varType.dtInterfaceInfo = classinfo.DefaultInterface
GetForm.tokVariable.varType.dtDataType = vbObject
Print #99, "11"
'Set GetForm.tokInterfaceInfo = GetForm.tokVariable.varType.dtInterfaceInfo
Print #99, "12"
' Form1(0)
'Set GetForm.tokReturnTypeInterfaceInfo = GetForm.tokVariable.varType.dtInterfaceInfo
'Set GetForm.tokMemberInfo = GetForm.tokVariable.varType.dtInterfaceInfo.GetMember(0)
Print #99, "13"
GetForm.tokString = m.Name ' was "_" & m.Name but not good for VBEmitter
Print #99, "14"
GetForm.tokVariable.varSymbol = GetForm.tokString
Print #99, "15"
Set GetForm.tokVariable.varModule = m
'Set GetForm.tokVariable.varProc = p ' form is module level - no proc
GetForm.tokType = tokVariable
Set GetForm.tokReference = tliVB
GetForm.tokDataType = vbObject
Set GetForm.tokModule = currentModule
Print #99, "16"
End Function

Function OptionalComma(ByVal tokens As Collection) As Boolean
If Not IsEOL(tokens) Then If tokens.Item(1).tokString = "," Then tokens.Remove 1: OptionalComma = True ' Else Err.Raise 1 - could be ") Else or :"
End Function

' get next batch of tokens, striping leading colons
Function GetNextTokenizedStatement(ByVal ln As Long, ByRef tokens As Collection, Optional removeall As Boolean) As Long
Print #99, "GetNextTokenizedStatement: ln="; ln; " t.c="; tokens.Count; " ra="; removeall
If removeall Then RemoveAllTokens tokens: GoTo 10 ' Sorry
If tokens.Count = 0 Then GoTo 20 ' Sorry
If tokens.Item(1).tokString <> ":" Then Err.Raise 1
Do
    tokens.Remove 1 ' :
10
    While tokens.Count = 0
20
        ln = GetNextTokenizedLine(ln, tokens)
    Wend
    If tokens.Item(1).tokString <> ":" Then Exit Do
Loop
GetNextTokenizedStatement = ln
Print #99, "GetNextTokenizedStatement: ln="; GetNextTokenizedStatement; "t.c="; tokens.Count; " t(1).s="; tokens.Item(1).tokString
End Function

Function GetNextTokenizedLine(ByVal ln As Long, ByRef tokens As Collection) As Long
Print #99, "GetNextTokenizedLine: ln="; ln; " cmt.c="; currentModule.ilines.Count
Set tokens = New Collection
If currentModule.ilines.Count <> 0 Then
    On Error Resume Next
    Set tokens = oTok.TokenizeLine(currentModule.lines.Item(CStr(ln)))
    On Error GoTo 0
    Dim token As vbToken
    For Each token In tokens
        Print #99, "tt="; token.tokType; " ts="; token.tokString
    Next
End If
GetNextTokenizedLine = ln + 1
Print #99, "GetNextTokenizedLine: t.c="; tokens.Count
End Function

Function IsForm(ByVal ct As vbext_ComponentType) As Boolean
Print #99, "IsForm: ct="; ct
Select Case ct
Case vbext_ct_MSForm, vbext_ct_VBForm, vbext_ct_VBMDIForm, vbext_ct_PropPage, vbext_ct_UserControl ', vbext_ct_ActiveXDesigner
    IsForm = True
Case Else
    ' do nothing
End Select
End Function

Function IsEOL(ByVal tokens As Collection, Optional ByVal i As Long = 1) As Boolean
If tokens.Count = 0 Then IsEOL = True Else If IsAny(UCase(tokens.Item(i).tokString), "ELSE", ":") Then IsEOL = True
End Function
