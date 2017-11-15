Public Function GetTrsEarningsDifferencesForRecalcLetter() As DataTable

        Dim DiffFullTable As DataTable = Nothing
        Dim TrsLetterDifferences As New DataTable
        Dim DataRow As DataRow = Nothing
        Dim AddDifference As Boolean = False
        Dim rowsAlreadyProcessed As New List(Of Integer)

        Try

            With TrsLetterDifferences
                .Columns.Add("FiscalYr", GetType(System.Int32))
                .Columns.Add("EarnStatusID", GetType(System.Int32))
                .Columns.Add("EarnSrceID", GetType(System.Int32))
                .Columns.Add("EmployTypeID", GetType(System.Int32))
                .Columns.Add("UseRate_Before", GetType(System.Decimal))
                .Columns.Add("UseRate_After", GetType(System.Decimal))
                .Columns.Add("ChangeTypeId", GetType(System.Int32))
            End With

            DiffFullTable = GetAllBeforeAndAfterRowDifferences(trsEarningsCompare.TableId.EarnRec)

            Dim JoinedDiffTable = From Before In DiffFullTable.Select("TypeId = 'Before'").CopyToDataTable.AsEnumerable
                                  Join After In DiffFullTable.Select("TypeId = 'After'").CopyToDataTable.AsEnumerable
                                    On Before.Field(Of Integer)("JoinId") Equals After.Field(Of Integer)("JoinId")


            For Each JoinedRow In JoinedDiffTable
                AddDifference = True
                Dim rowId As Integer? = 0

                If CType(Coalesce(JoinedRow.Before.Item("ChangeTypeId"), JoinedRow.After.Item("ChangeTypeId")), ChangeTypeCode) = ChangeTypeCode.ectcDelete AndAlso
                   CType(Coalesce(JoinedRow.Before.Item("EarnSrceId"), JoinedRow.After.Item("EarnSrceId")), EarnSourceCode) = EarnSourceCode.eeaSupplemental Then

                    rowId = (From row In DiffFullTable.AsEnumerable()
                             Let FiscalYrNullable = row.Field(Of Integer?)("FiscalYr"),
                             EmplyrIdNullable = row.Field(Of Integer?)("EmplyrId"),
                             EmployTypeIdNullable = row.Field(Of Byte?)("EmployTypeId"),
                             EarnSrceIdNullable = row.Field(Of Byte?)("EarnSrceId"),
                             ChangeTypeIdNullable = row.Field(Of Integer?)("ChangeTypeId"),
                             SalaryNullable = row.Field(Of Decimal?)("UseRate"),
                             Id = row.Field(Of Integer)("Id")
                             Where FiscalYrNullable.HasValue AndAlso FiscalYrNullable.Value = CInt(Coalesce(JoinedRow.Before.Item("FiscalYr"), JoinedRow.After.Item("FiscalYr"))) AndAlso
                             EmplyrIdNullable.HasValue AndAlso EmplyrIdNullable.Value = CInt(Coalesce(JoinedRow.Before.Item("EmplyrId"), JoinedRow.After.Item("EmplyrId"))) AndAlso
                             EmployTypeIdNullable.HasValue AndAlso EmployTypeIdNullable.Value = CInt(Coalesce(JoinedRow.Before.Item("EmployTypeId"), JoinedRow.After.Item("EmployTypeId"))) AndAlso
                             EarnSrceIdNullable.HasValue AndAlso CType(EarnSrceIdNullable.Value, EarnSourceCode) = EarnSourceCode.eeaAnnualReport AndAlso
                             SalaryNullable.HasValue AndAlso SalaryNullable.Value <> CDec(Coalesce(JoinedRow.Before.Item("UseRate"), JoinedRow.After.Item("UseRate"))) AndAlso
                             ChangeTypeIdNullable.HasValue AndAlso CType(ChangeTypeIdNullable.Value, ChangeTypeCode) = ChangeTypeCode.ectcAdd AndAlso
                             not rowsAlreadyProcessed.Contains(id) AndAlso
                             row.Field(Of String)("TypeId") = "After").FirstOrDefault()?.Id

                    AddDifference = rowId.GetValueOrDefault(0) > 0

                    If AddDifference Then
                        rowsAlreadyProcessed.Add(rowId.Value)
                        rowsAlreadyProcessed.Add(JoinedRow.After.Field(Of Integer)("Id"))
                        rowsAlreadyProcessed.Add(JoinedRow.Before.Field(Of Integer)("Id"))
                        Dim replacementRate = DiffFullTable.AsEnumerable().First(Function(row) row.Field(Of Integer)("Id") = rowId.Value)("UseRate")
                        rowsAlreadyProcessed.Add(rowId.Value)
                        If IsDBNullOrNothing(JoinedRow.Before.Item("UseRate")) Then
                            JoinedRow.Before.Item("UseRate") = replacementRate
                        Else
                            JoinedRow.After.Item("UseRate") = replacementRate
                        End If
                    End If


                Else
                    If CType(Coalesce(JoinedRow.Before.Item("ChangeTypeId"), JoinedRow.After.Item("ChangeTypeId")), ChangeTypeCode) = ChangeTypeCode.ectcAdd AndAlso
                       CType(Coalesce(JoinedRow.Before.Item("EarnSrceId"), JoinedRow.After.Item("EarnSrceId")), EarnSourceCode) = EarnSourceCode.eeaAnnualReport Then

                        rowId = (From row In DiffFullTable.AsEnumerable()
                                 Let FiscalYrNullable = row.Field(Of Integer?)("FiscalYr"),
                                 EmplyrIdNullable = row.Field(Of Integer?)("EmplyrId"),
                                 EmployTypeIdNullable = row.Field(Of Byte?)("EmployTypeId"),
                                 EarnSrceIdNullable = row.Field(Of Byte?)("EarnSrceId"),
                                 ChangeTypeIdNullable = row.Field(Of Integer?)("ChangeTypeId"),
                                 SalaryNullable = row.Field(Of Decimal?)("UseRate"),
                                 Id = row.Field(Of Integer)("Id")
                                 Where FiscalYrNullable.HasValue AndAlso FiscalYrNullable.Value = CInt(Coalesce(JoinedRow.Before.Item("FiscalYr"), JoinedRow.After.Item("FiscalYr"))) AndAlso
                                 EmplyrIdNullable.HasValue AndAlso EmplyrIdNullable.Value = CInt(Coalesce(JoinedRow.Before.Item("EmplyrId"), JoinedRow.After.Item("EmplyrId"))) AndAlso
                                 EmployTypeIdNullable.HasValue AndAlso EmployTypeIdNullable.Value = CInt(Coalesce(JoinedRow.Before.Item("EmployTypeId"), JoinedRow.After.Item("EmployTypeId"))) AndAlso
                                 EarnSrceIdNullable.HasValue AndAlso CType(EarnSrceIdNullable.Value, EarnSourceCode) = EarnSourceCode.eeaSupplemental AndAlso
                                 ChangeTypeIdNullable.HasValue AndAlso CType(ChangeTypeIdNullable.Value, ChangeTypeCode) = ChangeTypeCode.ectcDelete AndAlso
                                 SalaryNullable.HasValue AndAlso SalaryNullable.Value <> CDec(Coalesce(JoinedRow.Before.Item("UseRate"), JoinedRow.After.Item("UseRate"))) AndAlso
                                 row.Field(Of String)("TypeId") = "Before" AndAlso
                                 not rowsAlreadyProcessed.Contains(id) AndAlso
                                 Id <> rowId).FirstOrDefault()?.Id

                        AddDifference = rowId.GetValueOrDefault(0) > 0

                        If AddDifference Then
                            rowsAlreadyProcessed.Add(rowId.Value)
                            rowsAlreadyProcessed.Add(JoinedRow.After.Field(Of Integer)("Id"))
                            rowsAlreadyProcessed.Add(JoinedRow.Before.Field(Of Integer)("Id"))
                            Dim replacementRate = DiffFullTable.AsEnumerable().First(Function(row) row.Field(Of Integer)("Id") = rowId.Value)("UseRate")
                            If IsDBNullOrNothing(JoinedRow.Before.Item("UseRate")) Then
                                JoinedRow.Before.Item("UseRate") = replacementRate
                            Else
                                JoinedRow.After.Item("UseRate") = replacementRate
                            End If
                        End If

                    End If
                End If



                If AddDifference Then
                    DataRow = TrsLetterDifferences.NewRow
                    DataRow.Item("FiscalYr") = CInt(Coalesce(JoinedRow.Before.Item("FiscalYr"), JoinedRow.After.Item("FiscalYr")))
                    DataRow.Item("EarnStatusID") = CInt(Coalesce(JoinedRow.Before.Item("EarnStatusID"), JoinedRow.After.Item("EarnStatusID")))
                    DataRow.Item("EarnSrceID") = CInt(Coalesce(JoinedRow.Before.Item("EarnSrceID"), JoinedRow.After.Item("EarnSrceID")))
                    DataRow.Item("EmployTypeID") = CInt(Coalesce(JoinedRow.Before.Item("EmployTypeID"), JoinedRow.After.Item("EmployTypeID")))
                    DataRow.Item("UseRate_Before") = CDec(Coalesce(JoinedRow.Before.Item("UseRate"), 0))
                    DataRow.Item("UseRate_After") = CDec(Coalesce(JoinedRow.After.Item("UseRate"), 0))
                    DataRow.Item("ChangeTypeId") = Coalesce(JoinedRow.Before.Item("ChangeTypeId"), JoinedRow.After.Item("ChangeTypeId"))
                    
                    TrsLetterDifferences.Rows.Add(DataRow)
                End If

            Next

            Return TrsLetterDifferences

        Finally
            DiffFullTable?.Dispose()
            TrsLetterDifferences?.Dispose()
        End Try

    End Function