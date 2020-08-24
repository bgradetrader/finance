Attribute VB_Name = "Module2"
Option Explicit
Option Base 1

Public Function BlackScholes_Option_Price(call_put As String, s As Double, x As Double, t As Double, r As Double, b As Double, v As Double)


    
    Dim d1 As Double
    Dim d2 As Double
 
    Dim p As Double
    Dim value As Double
    Dim cp As Long
    

        
        d1 = (Log(s / x) + (b + (v ^ 2) / 2) * t) / (v * Sqr(t))
        d2 = d1 - v * Sqr(t)
        
        If call_put = "c" Or call_put = "C" Then
            
            BlackScholes_Option_Price = s * WorksheetFunction.NormSDist(d1) - x * Exp(b - r * t) * WorksheetFunction.NormSDist(d2)
            
        ElseIf call_put = "p" Or call_put = "P" Then
            
            BlackScholes_Option_Price = x * Exp(b - r * t) * WorksheetFunction.NormSDist(-d2) - s * WorksheetFunction.NormSDist(-d1)
            
        End If
        

    


End Function

Public Function BlackScholes_Option_Greeks(call_put As String, Output_flag As String, s As Double, x As Double, t As Double, r As Double, b As Double, v As Double)

    Dim ds As Double
    Dim dv As Double
    Dim dt As Double
    Dim dr As Double
    

    ds = s * 0.01
    dv = 0.01
    dt = 1 / 365
    dr = 0.01

    If Output_flag = "p" Or Output_flag = "P" Then
        BlackScholes_Option_Greeks = BlackScholes_Option_Price(call_put, s, x, t, r, b, v)
    ElseIf Output_flag = "d" Or Output_flag = "D" Then 'delta
        BlackScholes_Option_Greeks = (BlackScholes_Option_Price(call_put, s + ds, x, t, r, b, v) - BlackScholes_Option_Price(call_put, s - ds, x, t, r, b, v)) / (2 * ds)
    ElseIf Output_flag = "g" Or Output_flag = "G" Then 'gamma
        BlackScholes_Option_Greeks = (BlackScholes_Option_Price(call_put, s + ds, x, t, r, b, v) - 2 * BlackScholes_Option_Price(call_put, s, x, t, r, b, v) + BlackScholes_Option_Price(call_put, s - ds, x, t, r, b, v)) / (ds ^ 2) '�� ���� �����ٰ� �ְ� ��ȭ���� �����ָ� 1�ִ� �߰������� �Ÿ�������ϴ� ������ ����, �¼��� �ִٸ� �¼����� �����ָ� ��ü �ŷ������� ������ ����
    ElseIf Output_flag = "v" Or Output_flag = "V" Then 'vega
        BlackScholes_Option_Greeks = (BlackScholes_Option_Price(call_put, s, x, t, r, b, v + dv) - BlackScholes_Option_Price(call_put, s, x, t, r, b, v - dv)) / (2 * dv) '������ 1�� �������� �� 100%�� ����, 1%�� ������ ġȯ���ַ��� 100���� ������ ��
    ElseIf Output_flag = "t" Or Output_flag = "T" Then 'theta
        BlackScholes_Option_Greeks = BlackScholes_Option_Price(call_put, s, x, t - dt, r, b, v) - BlackScholes_Option_Price(call_put, s, x, t, r, b, v) '1�� ��Ÿ, �̰� ��ü�� ������ �ٸ��� �ΰ����ε� �̰Ŵ� amount����
    ElseIf Output_flag = "r" Or Output_flag = "R" Then 'Rho
        BlackScholes_Option_Greeks = (BlackScholes_Option_Price(call_put, s, x, t, r + dr, b, v) - BlackScholes_Option_Price(call_put, s, x, t, r - dr, b, v)) / (2 * dr) '������1(100%)�� rho ���� 100���� ������� 1%�� rho
    ElseIf Output_flag = "dddv" Or Output_flag = "dDdV" Then 'vanna
        BlackScholes_Option_Greeks = (BlackScholes_Option_Price(call_put, s + ds, x, t, r, b, v + dv) - BlackScholes_Option_Price(call_put, s - ds, x, t, r, b, v + dv) _
                                    - BlackScholes_Option_Price(call_put, s + ds, x, t, r, b, v - dv) + BlackScholes_Option_Price(call_put, s - ds, x, t, r, b, v - dv)) / (4 * ds * dv) ' ������1(100%) �� �ְ� 1 ���� vanna
    ElseIf Output_flag = "dddt" Or Output_flag = "dDdT" Then 'charm
        BlackScholes_Option_Greeks = (BlackScholes_Option_Price(call_put, s + ds, x, t + dt, r, b, v) - BlackScholes_Option_Price(call_put, s - ds, x, t + dt, r, b, v) _
                                    - BlackScholes_Option_Price(call_put, s + ds, x, t - dt, r, b, v) + BlackScholes_Option_Price(call_put, s - ds, x, t - dt, r, b, v)) / (2 * ds)
        
    End If
    
    



End Function

Public Function BlackScholes_ImpVol(call_put As String, s As Double, x As Double, t As Double, r As Double, b As Double, v As Double, Cprice As Double)

Dim e As Double
Dim Vlow, Vhigh, Vmid As Double
Dim Clow, Chigh, Cmid As Double
Dim i As Integer

    e = 0.000000001
    Vlow = 0.0000001
    Vhigh = 1
    
    Clow = BlackScholes_Option_Price(call_put, s, x, t, r, b, Vlow)
    Chigh = BlackScholes_Option_Price(call_put, s, x, t, r, b, Vhigh)
    
    If Cprice < Clow Then
        Vmid = Vlow
    ElseIf Cprice > Chigh Then
        Vmid = Vhigh
    Else
        For i = 1 To 100
            Vmid = (Vlow + Vhigh) / 2
            Cmid = BlackScholes_Option_Price(call_put, s, x, t, r, b, Vmid)
            
            If Abs(Cprice - Cmid) < e Then
                BlackScholes_ImpVol = Vmid
                Exit Function
            End If
            
            If Cmid < Cprice Then
                Vlow = Vmid
                Clow = Cmid
            Else
                Vhigh = Vmid
                Chigh = Cmid
            End If
        Next i
    End If
    
    BlackScholes_ImpVol = Vmid


End Function

Public Function Bi_Tree_Euro_Option_Price(call_put As String, Output_flag As String, stock_price As Double, strike_price As Double, remain_period As Double, interest_rate As Double, vol As Double)


    

End Function


Sub expl()

Call BlackScholes_Option_Greeks("c", "p", 45, 50, 0.3, 0.01, 0.4)


End Sub

