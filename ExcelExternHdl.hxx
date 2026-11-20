#ifndef _EXCEL_EXTERN_HDL_HXX_
#define _EXCEL_EXTERN_HDL_HXX_

#include <BaseExternHdl.hxx>

// Not "ExternHdl": libCtrlV321 exports a class of that name (the built-in
// CTRL functions), and on Linux the dynamic linker would bind this class's
// vtable and typeinfo to WinCC OA's, so none of the functions would exist.
class ExcelExternHdl : public BaseExternHdl
{
  public:
    ExcelExternHdl(BaseExternHdl *nextHdl, PVSSulong funcCount, FunctionListRec fnList[])
      : BaseExternHdl(nextHdl, funcCount, fnList) {}

    const Variable *execute(ExecuteParamRec &param) override;

    /// Assignable variable behind a reference argument (getTarget is
    /// protected); used again when an *Async result is delivered.
    Variable *resolveTarget(CtrlExpr *expr, const ExecuteParamRec &param) const
    {
      return getTarget(expr, param, NO_VAR);
    }
};

#endif
