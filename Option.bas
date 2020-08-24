Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Sub path_dependent_option_binomial() 'for call option
Dim stock(), strike_value(), backwardation_value()
Dim i, j, m As Integer


    Dim s, k, vol, div, ir, step As Double            'input 변수: function으로 변경 가능
    Dim pricedate, matdate As Date                    'input 변수: function으로 변경 가능
    
    Dim dt, u, d, a, p, p1, disf, rf, t As Double

    pricedate = "2019-02-28"                           '평가일
    matdate = "2020-02-28"                             '만기일
    ir = 0.1                                           '이자율
    vol = 0.4                                          '변동성
    s = 100                                            '현재주가
    div = 0.08                                         '정률배당
    k = 60                                             '행사가
    step = 6                                           '구간개수
    
    t = WorksheetFunction.YearFrac(pricedate, matdate)
    rf = Log(1 + ir)
    dt = t / step
    u = Exp(vol * Sqr(dt))
    d = 1 / u
    a = Exp((rf - div) * dt)
    p = (a - d) / (u - d)
    p1 = 1 - p
    disf = Exp(-rf * dt)
    
    ReDim stock(step + 1, step + 1) '주가에 대한 트리를 구하는 루프
    
        stock(1, 1) = s
        
        For i = 2 To step + 1
            
            stock(1, i) = stock(1, i - 1) * u
            
                For j = 2 To i
                        
                    stock(j, i) = stock(j - 1, i) / u * d
                    
                Next j
                
        Next i
    
    ReDim strike_value(step + 1, step + 1) '각 노드에서 옵션을 행사하였을때 얻을 수 있는 가치를 구하는 루프
    
        For i = 2 To step + 1
        
            For j = 1 To i
            
                strike_value(j, i) = WorksheetFunction.Max(stock(j, i) - k, 0)
                
            Next j
 
        Next i
        
    ReDim backwardation_value(step + 1, step + 1) '마지막 지점에서 이전노드로 backwardation을 수행
                                                  '마지막에서 두번째 노드부터 첫 노드까지는 backwardation을 수행하여 구한 가치와 행사하였을때 얻을 수 있는 가치를 비교하여
                                                  '큰값을 가져오는 작업을 수행
    
        For i = 1 To step + 1
            
            backwardation_value(i, step + 1) = strike_value(i, step + 1)
        
        Next i
        
        For i = 1 To step
        
            j = step + 1 - i
            
            For m = 1 To j
            
                backwardation_value(m, j) = disf * (p * backwardation_value(m, j + 1) + p1 * backwardation_value(m + 1, j + 1))
                
                If backwardation_value(m, j) <= strike_value(m, j) Then
                   backwardation_value(m, j) = strike_value(m, j)
                End If
                
            Next m
            
        Next i
        


        MsgBox backwardation_value(1, 1) '아메리칸 옵션의 가치
                
''''''노드값 테스트'''''''
Worksheets("Sheet1").Range(Worksheets("sheet1").Cells(1, 1), Worksheets("sheet1").Cells(1 + step, 1 + step)) = stock()
Worksheets("Sheet2").Range(Worksheets("sheet2").Cells(1, 1), Worksheets("sheet2").Cells(1 + step, 1 + step)) = strike_value()
Worksheets("Sheet3").Range(Worksheets("sheet3").Cells(1, 1), Worksheets("sheet3").Cells(1 + step, 1 + step)) = backwardation_value()


End Sub

