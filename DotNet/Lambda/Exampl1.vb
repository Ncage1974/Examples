 'Join Syntax
 ' .Join(JoinObject,
        'Left Key To Join On,
        'Right Key to Join On,
        'Data that will come out of the join)
 
 Dim commonFuncs As New CommonFunctionsLite
            Using claimEarningsObjBefore = New ClaimEarningsObj(TrsApp.ADO, ClaimRecalcId, False, True)
                Using claimEarningsObjAfter = New ClaimEarningsObj(TrsApp.ADO, ClaimRecalcId, True, True)
                    Dim earningsBefore = claimEarningsObjBefore.EarningsList.AsEnumerable()
                    Dim earningsAfter = claimEarningsObjAfter.EarningsList.AsEnumerable()
                    Dim earningsQuery =
                     earningsBefore.
                     Join(earningsAfter,
                          Function(beforeKey) New With {Key .FiscalYr = beforeKey(EarningsCols.FiscalYr), Key .EmpId = beforeKey(EarningsCols.EmplyrId)},
                          Function(afterKey) New With {Key .FiscalYr = afterKey(EarningsCols.FiscalYr), Key .EmpId = afterKey(EarningsCols.EmplyrId)},
                          Function(before, after)
                              Return New With {
                                 .FiscalYear = DirectCast(before(EarningsCols.FiscalYr), Integer),
                                 .OriginalSalary = DirectCast(before(EarningsCols.Rate), Decimal),
                                 .RevisedSalary = DirectCast(after(EarningsCols.Rate), Decimal),
                                 .RecipId = commonFuncs.InferCoalesce(before(EarningsCols.RecipMatchDtlId), after(EarningsCols.RecipMatchDtlId))
                              }
                          End Function).
                      Where(Function(beforeAndAfter)
                                Return IsNothing(beforeAndAfter.RecipId) AndAlso
                                  beforeAndAfter.OriginalSalary <> beforeAndAfter.RevisedSalary
                            End Function).
                      Select(Function(beforeAndAfter)
                                 Return New SalaryDifferences(beforeAndAfter.FiscalYear, beforeAndAfter.OriginalSalary, beforeAndAfter.RevisedSalary)
                             End Function).
                      ToList()