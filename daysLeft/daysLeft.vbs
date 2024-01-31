' 終了年月日、有休取得日数を入力値で取得
' 固定値であれば第三引数に入力する
getEndDate = InputBox("終了年月日を入力してください","入力")　
getPaidHoliday = InputBox("有休取得日数を入力してください","入力")

startDate = Date()
endDate = getEndDate
saturdayCnt = GetWeekDay(startDate,endDate,7)
sundayCnt = GetWeekDay(startDate,endDate,1)
' 祝日に関しては入力しないと難しい
holidayCnt = 1
paidHoliday = getPaidHoliday

' 開始年月日から終了年月日までの日数を取得
dateCnt = DateDiff("d",startDate,endDate)

' 開始年月日から終了年月日の日数 - 土曜の日数 - 日曜の日数 - 祝日の日数 - 有休の日数
result = dateCnt - saturdayCnt - sundayCnt - holidayCnt - paidHoliday

' メッセージ表示
msg = "残り日数は、"&result
MsgBox (msg)


' 指定期間の指定曜日の日数を取得するメソッド
' dtStart   開始年月日
' dtEnd     終了年月日
' iWeekDay  取得曜日    1=日曜日、2=月曜日、3=火曜日、4=水曜日、5=木曜日、6=金曜日、7=土曜日
Public Function GetWeekDay(dtStart,dtEnd,iWeekDay)

    '計算処理
    ' 期間の日数を取得
    lngDay = DateDiff("d",dtStart,dtEnd) + 1

    ' 週の数を取得
    lngWeek = Int(lngDay / 7)

    ' 余った日数を取得
    lngRest = lngDay - lngWeek * 7

    ' 終了日の曜日を取得
    lngEndDay = WeekDay(dtEnd,iWeekDay)

    ' 条件判断
    If lngEndDay > lngRest Then
        GetWeekDay = lngWeek
    Else
        GetWeekDay = lngWeek + 1
    End IF

End Function