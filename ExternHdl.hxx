#ifndef _EXTERNHDL_H_
#define _EXTERNHDL_H_

#include <BaseExternHdl.hxx>

class ExternHdl : public BaseExternHdl
{
  public:
    ExternHdl(BaseExternHdl *nextHdl, PVSSulong funcCount, FunctionListRec fnList[])
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
