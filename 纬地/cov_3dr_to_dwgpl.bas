Attribute VB_Name = "ģ��1"
Sub cov_3dr_to_dwgpl()
'���ܣ���3dr����תΪcad�ж������������ʽ�������������߶Ա�
sheetname = "sheet1" '�Ȱ�3dr�����ݸ��Ƶ�sheetname�У���ɾ����һ��
row_max = 463
output_sheetname = "sheet2"

    For i = 1 To row_max
        Change = Sheets(sheetname).Cells(i, 1)
        If Sheets(sheetname).Cells(i, 2) = "" And Change <> "" Then
            
            Sheets(output_sheetname).Cells(i, 1) = Change
            For j = 1 To 2
                row_data = i + j
                pline_string = ""
                col_max = Sheets(sheetname).Cells(row_data, 1)
                h_absolute = Sheets(sheetname).Cells(row_data, 2)
                v_absolute = Sheets(sheetname).Cells(row_data, 3)
                v_relative = 0
                h_relative = 0
                For i_col = 1 To col_max
                    On Error Resume Next
                    h_relative = Sheets(sheetname).Cells(row_data, i_col * 2) - Sheets(sheetname).Cells(row_data, i_col * 2 - 2)
                    If i_col = 1 Then
                        v_relative = Sheets(sheetname).Cells(row_data, i_col * 2 + 1) - v_absolute
                    Else
                        v_relative = Sheets(sheetname).Cells(row_data, i_col * 2 + 1) - Sheets(sheetname).Cells(row_data, i_col * 2 - 1)
                    End If
                    pline_string = pline_string & "@" & h_relative & "," & v_relative & " "
                Next i_col
                Sheets(output_sheetname).Cells(row_data, 1) = pline_string
            Next j
    '        col_right_max = Sheets(sheetname).Cells(i + 2, 1)
        End If
        
    Next i

End Sub
