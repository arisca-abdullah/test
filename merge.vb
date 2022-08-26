Sub MergeCell()
    Dim required_input As String
    Dim required_data as Variant

    required_input = InputBox("Masukkan posisi kolom check sheet, unit, plant, disiplin ilmu, baris awal, dan baris akhir (Contoh: A,B,C,D,1,100)", "Input data")
    required_data = Split(required_input, ",")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim csStart As Integer
    Dim prevCs As String
    Dim currCs As String

    Dim unitStart As Integer
    Dim prevUnit As String
    Dim currUnit As String

    Dim plantStart As Integer
    Dim prevPlant As String
    Dim currPlant As String

    Dim disiplinIlmuStart As Integer
    Dim prevDisiplinIlmu As String
    Dim currDisiplinIlmu As String

    For i = required_data(4) To (required_data(5) + 1)
        If i = (required_data(4) + 0) Then
            csStart = i
            prevCs = Range(required_data(0) & i).Value
            currCs = Range(required_data(0) & i).Value

            unitStart = i
            prevUnit = Range(required_data(1) & i).Value
            currUnit = Range(required_data(1) & i).Value

            plantStart = i
            prevPlant = Range(required_data(2) & i).Value
            currPlant = Range(required_data(2) & i).Value

            disiplinIlmuStart = i
            prevDisiplinIlmu = Range(required_data(3) & i).Value
            currDisiplinIlmu = Range(required_data(3) & i).Value
        ElseIf i <= required_data(5) Then
            currCs = Range(required_data(0) & i).Value
            currUnit = Range(required_data(1) & i).Value
            currPlant = Range(required_data(2) & i).Value
            currDisiplinIlmu = Range(required_data(3) & i).Value

            If currCs = "" Then
                currCs = prevCs
            End If

            If currUnit = "" Then
                currUnit = prevUnit
            End If

            If currPlant = "" Then
                currPlant = prevPlant
            End If

            If currDisiplinIlmu = "" Then
                currDisiplinIlmu = prevDisicurrDisiplinIlmu
            End If

            If prevCs <> currCs Then
                If csStart <> (i - 1) Then
                    Range(required_data(0) & csStart & ":" & required_data(0) & (i - 1)).Merge
                End If

                If unitStart <> (i - 1) Then
                    Range(required_data(1) & unitStart & ":" & required_data(1) & (i - 1)).Merge
                End If

                If plantStart <> (i - 1) Then
                    Range(required_data(2) & plantStart & ":" & required_data(2) & (i - 1)).Merge
                End If

                If disiplinIlmuStart <> (i - 1) Then
                    Range(required_data(3) & disiplinIlmuStart & ":" & required_data(3) & (i - 1)).Merge
                End If

                csStart = i
                prevCs = currCs

                unitStart = i
                prevUnit = currUnit

                plantStart = i
                prevPlant = currPlant

                disiplinIlmuStart = i
                prevDisiplinIlmu = currDisiplinIlmu
            ElseIf prevUnit <> currUnit Then
                If unitStart <> (i - 1) Then
                    Range(required_data(1) & unitStart & ":" & required_data(1) & (i - 1)).Merge
                End If

                If plantStart <> (i - 1) Then
                    Range(required_data(2) & plantStart & ":" & required_data(2) & (i - 1)).Merge
                End If

                If disiplinIlmuStart <> (i - 1) Then
                    Range(required_data(3) & disiplinIlmuStart & ":" & required_data(3) & (i - 1)).Merge
                End If

                unitStart = i
                prevUnit = currUnit

                plantStart = i
                prevPlant = currPlant

                disiplinIlmuStart = i
                prevDisiplinIlmu = currDisiplinIlmu
            ElseIf prevPlant <> currPlant Then
                If plantStart <> (i - 1) Then
                    Range(required_data(2) & plantStart & ":" & required_data(2) & (i - 1)).Merge
                End If

                If disiplinIlmuStart <> (i - 1) Then
                    Range(required_data(3) & disiplinIlmuStart & ":" & required_data(3) & (i - 1)).Merge
                End If

                plantStart = i
                prevPlant = currPlant

                disiplinIlmuStart = i
                prevDisiplinIlmu = currDisiplinIlmu
            ElseIf prevDisiplinIlmu <> currDisiplinIlmu Then
                If disiplinIlmuStart <> (i - 1) Then
                    Range(required_data(3) & disiplinIlmuStart & ":" & required_data(3) & (i - 1)).Merge
                End If

                disiplinIlmuStart = i
                prevDisiplinIlmu = currDisiplinIlmu
            End If
        Else
            If csStart <> (i - 1) Then
                Range(required_data(0) & csStart & ":" & required_data(0) & (i - 1)).Merge
            End If

            If unitStart <> (i - 1) Then
                Range(required_data(1) & unitStart & ":" & required_data(1) & (i - 1)).Merge
            End If

            If plantStart <> (i - 1) Then
                Range(required_data(2) & plantStart & ":" & required_data(2) & (i - 1)).Merge
            End If

            If disiplinIlmuStart <> (i - 1) Then
                Range(required_data(3) & disiplinIlmuStart & ":" & required_data(3) & (i - 1)).Merge
            End If
        End If
    Next

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
