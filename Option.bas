Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Sub path_dependent_option_binomial() 'for call option
Dim stock(), strike_value(), backwardation_value()
Dim i, j, m As Integer


    Dim s, k, vol, div, ir, step As Double            'input ����: function���� ���� ����
    Dim pricedate, matdate As Date                    'input ����: function���� ���� ����
    
    Dim dt, u, d, a, p, p1, disf, rf, t As Double

    pricedate = "2019-02-28"                           '����
    matdate = "2020-02-28"                             '������
    ir = 0.1                                           '������
    vol = 0.4                                          '������
    s = 100                                            '�����ְ�
    div = 0.08                                         '�������
    k = 60                                             '��簡
    step = 6                                           '��������
    
    t = WorksheetFunction.YearFrac(pricedate, matdate)
    rf = Log(1 + ir)
    dt = t / step
    u = Exp(vol * Sqr(dt))
    d = 1 / u
    a = Exp((rf - div) * dt)
    p = (a - d) / (u - d)
    p1 = 1 - p
    disf = Exp(-rf * dt)
    
    ReDim stock(step + 1, step + 1) '�ְ��� ���� Ʈ���� ���ϴ� ����
    
        stock(1, 1) = s
        
        For i = 2 To step + 1
            
            stock(1, i) = stock(1, i - 1) * u
            
                For j = 2 To i
                        
                    stock(j, i) = stock(j - 1, i) / u * d
                    
                Next j
                
        Next i
    
    ReDim strike_value(step + 1, step + 1) '�� ��忡�� �ɼ��� ����Ͽ����� ���� �� �ִ� ��ġ�� ���ϴ� ����
    
        For i = 2 To step + 1
        
            For j = 1 To i
            
                strike_value(j, i) = WorksheetFunction.Max(stock(j, i) - k, 0)
                
            Next j
 
        Next i
        
    ReDim backwardation_value(step + 1, step + 1) '������ �������� �������� backwardation�� ����
                                                  '���������� �ι�° ������ ù �������� backwardation�� �����Ͽ� ���� ��ġ�� ����Ͽ����� ���� �� �ִ� ��ġ�� ���Ͽ�
                                                  'ū���� �������� �۾��� ����
    
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
        


        MsgBox backwardation_value(1, 1) '�Ƹ޸�ĭ �ɼ��� ��ġ
                
''''''��尪 �׽�Ʈ'''''''
Worksheets("Sheet1").Range(Worksheets("sheet1").Cells(1, 1), Worksheets("sheet1").Cells(1 + step, 1 + step)) = stock()
Worksheets("Sheet2").Range(Worksheets("sheet2").Cells(1, 1), Worksheets("sheet2").Cells(1 + step, 1 + step)) = strike_value()
Worksheets("Sheet3").Range(Worksheets("sheet3").Cells(1, 1), Worksheets("sheet3").Cells(1 + step, 1 + step)) = backwardation_value()


End Sub

