Option Explicit

Dim h, j, k, l As Integer, Moodys_Cl_A_Rating As String, Moodys_Cl_B_Rating As String
Dim SP_Cl_A_Rating As String, SP_Cl_B_Rating As String, Monitoring As Boolean
Dim Cl_A_Yield_Reduction As Double, Cl_B_Yield_Reduction As Double
Dim Cl_A_Default_Rate As Double, Cl_B_Default_Rate As Double
Dim Monte_Carlo As Integer, Monte_Carlo_Runs As Integer
Dim Loss_Rate As Double, Rating_Scale_Array_M As Variant, Convergence_Criterion As Double
Dim Rating_Scale_Array_S As Variant, Class_A_Relaxation_Factor As Double, Class_B_Relaxation_Factor As Double
Dim Start_Time As Double, End_Time As Double, Elapsed_Time As Double, Non_Linear_Limit As Integer
Dim r_infinity As Double, alfa_parm As Double, beta_parm As Double, delta_parm
Dim Delta_A As Double, Delta_B As Double, Yield_A As Double, Yield_B As Double, WAL_A As Double, WAL_B As Double
Dim WAL_A_Index As Integer, WAL_B_Index As Integer, Non_Linear_Index As Integer
Dim New_Class_A_Rate As Double, New_Class_B_Rate As Double, Epsilon As Double
Dim IRRa As Double, IRRb As Double, DIRRa As Double, DIRRb As Double


'Asser Part Parameter Definition
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Term, Ships, UsefulLife As Integer
Public MarketValue As Double, ProbofSold As Double, ExpirationDate As Date
Public LengthUnit As String, Rate As Double, Opex As Double, SurExp As Double
Public StartDate As Date, StartT() As Integer
Public Lease_Count() As Integer, Lease() As Integer, CharterLength() As Integer
Public ScrapValue() As Double, YearBuilt() As Integer, TodayAge() As Integer
Public ReleaseInterval_Countdown() As Integer, Capacity() As Integer, Age() As Integer
Public ScrapIncome() As Double, Revenue() As Double, Total_Revenue() As Double
Public Cost() As Double, Total_Cost() As Double, Amort() As Double
Public Total_Amort() As Double, Total_ScrapIncome() As Double, Total_Sales() As Double
Public IsSold() As Boolean, sValue() As Double, HasBeenScrap() As Boolean
Public TodayAge_initial(), Capacity_initial(), ScrapValue_initial(), sValue_initial() As Variant
Public CharterLength_initial(), Rate_initial(), Opex_initial(), SurExp_initial(), ExpirationDate_initial() As Variant

Public ExecutedOnce() As Boolean
Public t As Integer 'Month 1 to 360
Public i As Integer 'No. 1 to 25 ship
Public a As Integer, B As Integer, ReleaseInterval As Integer 'releasing intervals, uniform distribution
Public UninflatedRate() As Double, Y As Integer, M As Integer
Public InflationRate As Double, V As Double, WAI As Double
Public SInflationFactor(1 To 25) As Double
Public InflationAnnFactor(1 To 50) As Double
Public InflationMonFactor(1 To 50) As Double

'Liability Part Variable Definition
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public MRperShip As Double 'maintenance reserve per ship per year
Public RAratio As Double 'Interest reserve account ratio
Public IPB As Double 'initial pool balance
Public WAImonth As Integer 'target number of months of WAI in reserve account
Public IR As Double 'eligible investment rate (APR)
Public DR As Double
Public initial_maintenance_reserve As Double, initial_interest_reserve As Double
Public Cl_A_Rate As Double, Cl_B_Rate As Double, Cl_C_Rate As Double
Public Cl_A_Size As Double, Cl_B_Size As Double, Cl_C_Size As Double
Public Cl_A_Bal As Double, Cl_B_Bal As Double, Cl_C_Bal As Double
Public Cl_A_M As Double, Cl_B_M As Double, Cl_C_M As Double
Public Cl_A_MP As Double, Cl_B_MP As Double, q As Integer
Public aM As Integer, bM As Integer
Public F0(), F1(), F2(), F3(), F4(), F5(), F6(), F7(), F8(), F9(), F10(), F11(), F12() As Double
Public N As Integer, RdM As Double, RdR As Double
Public RrM As Double, RaM() As Double, RS() As Double
Public PA() As Double, PAP() As Double, PAS() As Double
Public PB(), PBP() As Double, PBS() As Double
Public PC() As Double, PCP() As Double, PCS() As Double
Public ALa As Double, ALb As Double, IRR As Double

Public CurrentCollection() As Double
Public BalA() As Double, BalB() As Double, BalC() As Double
Public BegMR() As Double, BegRA() As Double
Public s(), SP(), SS() As Double
Public IA() As Double, IAP() As Double, IAS() As Double
Public IB() As Double, IBP() As Double, IBS() As Double
Public EndMR(), EndRA() As Double, RrR() As Double, RaR() As Double
Public IC() As Double, ICP() As Double, ICS() As Double
Public CICP() As Double, Fa() As Double, Fb() As Double
Public AAmort() As Double, BAmort() As Double
Public Month() As Integer, TACa() As Double, TACb() As Double
Public PMTa As Range, PMTb As Range, PMTc As Range

Public ct As Integer
Public ScrapInf_Lower As Double, ScrapInf_Upper, ScrapInf_alpha, ScrapInf_beta As Double
Public Inf_Lower, Inf_Upper, Inf_alpha, Inf_beta As Double
Public AnnualInflationRate, product, Sale_prob, New_A, New_B As Double
Public Opex_Lower, Opex_Upper, Opex_alpha, Opex_beta As Double
Public BasicOpex, AddOpex As Double, TotalOpex As Double
Public myTableArray201, myTableArray202, myTableArray141, myTableArray142 As Range
Public myTableArray091, myTableArray092, myTableArray01, myTableArray02 As Range
Public Sales1_Lower, Sales1_Upper, Sales1_alpha, Sales1_beta As Double
Public Sales2_Lower, Sales2_Upper, Sales2_alpha, Sales2_beta As Double
Public Sales3_Lower, Sales3_Upper, Sales3_alpha, Sales3_beta As Double
Public Sales4_Lower, Sales4_Upper, Sales4_alpha, Sales4_beta As Double


Function Yield_Curve(Time As Double, IRR_Drop As Double)
'
' Yield Curve Model = Treasury Curve + Credit Spread (this simple model does not include a liquidity premium)
' A liquidity spread can be trivially added in.
'
Yield_Curve = r_infinity / (1 + beta_parm * Exp(-delta_parm * (Time / 12))) + alfa_parm * Sqr((Time / 12) * 0.01 * IRR_Drop)

End Function



Sub main()

Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False
Application.ScreenUpdating = False
Application.EnableEvents = False

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Worksheets("Liability").Cells(7, 5).Value = 0
Worksheets("Liability").Cells(9, 5).Value = 0
Start_Time = Now()

Monte_Carlo_Runs = Worksheets("Liability").Range("Monte_Carlo_Runs").Value

Monitoring = Worksheets("Liability").Range("Monitor").Value
Randomize
'
' Set Up the Non-Linear Convergence Algorithm
'
r_infinity = Worksheets("Liability").Range("Yield_Curve_r_infinity").Value
alfa_parm = Worksheets("Liability").Range("Yield_Curve_alfa_parm").Value
beta_parm = Worksheets("Liability").Range("Yield_Curve_beta_parm").Value
delta_parm = Worksheets("Liability").Range("Yield_Curve_delta_parm").Value
Class_A_Relaxation_Factor = Worksheets("Liability").Range("Class_A_Relaxation").Value
Class_B_Relaxation_Factor = Worksheets("Liability").Range("Class_B_Relaxation").Value
Non_Linear_Limit = Worksheets("Liability").Range("Convergence_Limit").Value

Non_Linear_Index = 0
Epsilon = Worksheets("Liability").Range("Epsilon").Value

Worksheets("Liability").Range("Convergence_Status").Value = "Iteration 0"


'table setup''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set the Moody's rating scale array
With Worksheets("Graphs")
Set Rating_Scale_Array_M = .Range(.Cells(3, 9), .Cells(20, 10))
End With

' Set the S&P rating scale array
With Worksheets("Graphs")
Set Rating_Scale_Array_S = .Range(.Cells(3, 7), .Cells(20, 8))
End With

' Set ship property vlookup table 20
With Worksheets("LR20+")
Set myTableArray201 = .Range(.Cells(2, 3), .Cells(467, 4))
Set myTableArray202 = .Range(.Cells(2, 4), .Cells(467, 10))
End With

' Set ship property vlookup table 14-19
With Worksheets("LR14-19")
Set myTableArray141 = .Range(.Cells(2, 3), .Cells(931, 4))
Set myTableArray142 = .Range(.Cells(2, 4), .Cells(931, 10))
End With
                                              
' Set ship property vlookup table 9-13
With Worksheets("LR9-13")
Set myTableArray091 = .Range(.Cells(2, 3), .Cells(1933, 4))
Set myTableArray092 = .Range(.Cells(2, 4), .Cells(1933, 10))
End With

' Set ship property vlookup table 0-8
With Worksheets("LR0-8")
Set myTableArray01 = .Range(.Cells(2, 3), .Cells(688, 4))
Set myTableArray02 = .Range(.Cells(2, 4), .Cells(688, 10))
End With


'Constent initialize
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Term = Worksheets("Data").Range("Term").Value
Ships = Worksheets("Data").Range("Ships").Value
UsefulLife = Worksheets("Data").Range("UsefulLife ").Value
Sale_prob = Worksheets("Data").Range("Sales").Value

'scrapvalue inflation constent parameters
a = Worksheets("Data").Range("Uniform_a").Value
B = Worksheets("Data").Range("Uniform_b").Value
ScrapInf_Lower = Worksheets("Data").Range("ScrapInf_L").Value
ScrapInf_Upper = Worksheets("Data").Range("ScrapInf_U").Value
ScrapInf_alpha = Worksheets("Data").Range("ScrapInf_alpha").Value
ScrapInf_beta = Worksheets("Data").Range("ScrapInf_beta").Value

'rates inflation constent parameters
Inf_Lower = Worksheets("Data").Range("Inf_Lower").Value
Inf_Upper = Worksheets("Data").Range("Inf_Upper").Value
Inf_alpha = Worksheets("Data").Range("Inf_alpha").Value
Inf_beta = Worksheets("Data").Range("Inf_beta").Value

'operating expense constent parameters
Opex_Lower = Worksheets("Data").Range("Opex_L").Value
Opex_Upper = Worksheets("Data").Range("Opex_U").Value
Opex_alpha = Worksheets("Data").Range("Opex_alpha").Value
Opex_beta = Worksheets("Data").Range("Opex_beta").Value

'impromptusales constent parameters
'input parameters of beta distributions in terms of four buckets
Sales1_Lower = Worksheets("Data").Range("Sales1_L").Value
Sales1_Upper = Worksheets("Data").Range("Sales1_U").Value
Sales1_alpha = Worksheets("Data").Range("Sales1_alpha").Value
Sales1_beta = Worksheets("Data").Range("Sales1_beta").Value
Sales2_Lower = Worksheets("Data").Range("Sales2_L").Value
Sales2_Upper = Worksheets("Data").Range("Sales2_U").Value
Sales2_alpha = Worksheets("Data").Range("Sales2_alpha").Value
Sales2_beta = Worksheets("Data").Range("Sales2_beta").Value
Sales3_Lower = Worksheets("Data").Range("Sales3_L").Value
Sales3_Upper = Worksheets("Data").Range("Sales3_U").Value
Sales3_alpha = Worksheets("Data").Range("Sales3_alpha").Value
Sales3_beta = Worksheets("Data").Range("Sales3_beta").Value
Sales4_Lower = Worksheets("Data").Range("Sales4_L").Value
Sales4_Upper = Worksheets("Data").Range("Sales4_U").Value
Sales4_alpha = Worksheets("Data").Range("Sales4_alpha").Value
Sales4_beta = Worksheets("Data").Range("Sales4_beta").Value

V = 0


''''''''''''''''''''''initialize global variables''''''''''''''''''''''''

ReDim YearBuilt(Ships)
ReDim Amort(Ships)
ReDim TodayAge_initial(Ships)
ReDim Capacity_initial(Ships)
ReDim ScrapValue_initial(Ships)
ReDim sValue_initial(Ships)
ReDim CharterLength_initial(Ships)
ReDim Rate_initial(Ships)
ReDim Opex_initial(Ships)
ReDim SurExp_initial(Ships)
ReDim ExpirationDate_initial(Ships)

'Ship market value''''''''''''''''''''''''''''''''''''''''''''
For i = 1 To Ships

        MarketValue = Worksheets("Data").Cells(3 + i, 9).Value
        'calculate initial cost of the ships
        V = V + MarketValue * 1000000 'in dollars
        YearBuilt(i) = Worksheets("Data").Cells(3 + i, 2).Value
        Amort(i) = (MarketValue * 1000000) / ((UsefulLife - (2019 - YearBuilt(i))) * 12) 'Amortization of every ship each month
        
        TodayAge_initial(i) = Worksheets("Data").Cells(3 + i, 3).Value
        Capacity_initial(i) = Worksheets("Data").Cells(3 + i, 5).Value
        ScrapValue_initial(i) = Worksheets("Data").Cells(3 + i, 10).Value
        sValue_initial(i) = Worksheets("Data").Cells(3 + i, 10).Value
        CharterLength_initial(i) = Worksheets("Data").Cells(3 + i, 12).Value
        Rate_initial(i) = Worksheets("Data").Cells(3 + i, 14).Value
        Opex_initial(i) = Worksheets("Data").Cells(3 + i, 15).Value
        SurExp_initial(i) = Worksheets("Data").Cells(3 + i, 16).Value
        ExpirationDate_initial(i) = Worksheets("Data").Cells(3 + i, 11).Value

Next i


''''''''''''''''''''''constent parameters''''''''''''''''''''''''
MRperShip = Worksheets("Liability").Range("Maintenance_Reserve").Value 'maintenance reserve per ship per year
RAratio = Worksheets("Liability").Range("Reserve_Account").Value 'Interest reserve account ratio
'calculate Initial Pool Balance
IPB = V + MRperShip * Ships + RAratio * V 'initial pool balance

WAImonth = Worksheets("Liability").Range("n").Value 'target number of months of WAI in reserve account
WAI = Worksheets("Liability").Range("WAI").Value
IR = Worksheets("Liability").Range("Investment_Rate").Value 'eligible investment rate (APR)
DR = Worksheets("Liability").Range("D_Rate").Value

Cl_A_Rate = Worksheets("Liability").Range("Class_A_Rate").Value
Cl_B_Rate = Worksheets("Liability").Range("Class_B_Rate").Value
New_A = Cl_A_Rate
New_B = Cl_B_Rate
Cl_C_Rate = Worksheets("Liability").Range("Class_C_Rate").Value

Cl_A_Size = Worksheets("Liability").Range("Class_A_Size").Value
Cl_B_Size = Worksheets("Liability").Range("Class_B_Size").Value
Cl_C_Size = Worksheets("Liability").Range("Class_C_Size").Value

Cl_A_Bal = IPB * Cl_A_Size
Cl_B_Bal = IPB * Cl_B_Size
Cl_C_Bal = IPB * Cl_C_Size

Cl_A_M = Worksheets("Liability").Range("Class_A_M").Value
Cl_B_M = Worksheets("Liability").Range("Class_B_M").Value
Cl_C_M = Worksheets("Liability").Range("Class_C_M").Value

Cl_A_MP = Cl_A_Bal * (Cl_A_Rate / 12) / (1 - (1 + Cl_A_Rate / 12) ^ (-Cl_A_M))
Cl_B_MP = Cl_B_Bal * (Cl_B_Rate / 12) / (1 - (1 + Cl_B_Rate / 12) ^ (-Cl_B_M))

'TAC schedule''''''''''''''''''''''''''''''''''''''''''''
ReDim TACa(Term), TACb(Term)

For k = 0 To Term - 1

TACa(k) = 12 * Cl_A_MP / Cl_A_Rate * (1 - (1 + Cl_A_Rate / 12) ^ (k - Cl_A_M))
If TACa(k) <= 0 Then TACa(k) = 0

TACb(k) = 12 * Cl_B_MP / Cl_B_Rate * (1 - (1 + Cl_B_Rate / 12) ^ (k - Cl_B_M))
If TACb(k) <= 0 Then TACb(k) = 0

Next k

'Preliminaries''''''''''''''''''''''''''''''''''''''''''

initial_maintenance_reserve = MRperShip * Ships
initial_interest_reserve = RAratio * V




''''''''''''''''''''''''''Non-Linear Loop (Banach's Fixed Point Theorem)''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Do
Non_Linear_Index = Non_Linear_Index + 1
Cl_A_Yield_Reduction = 0: Cl_A_Default_Rate = 0
Cl_B_Yield_Reduction = 0: Cl_B_Default_Rate = 0
WAL_A_Index = 0: WAL_B_Index = 0
WAL_A = 0: WAL_B = 0


''''''''''''''''''''''''''''' Linear Loop (Monte Carlo Simulation)'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For h = 1 To Monte_Carlo_Runs


Call one_monte





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Cl_A_Yield_Reduction = Cl_A_Yield_Reduction + 10000 * DIRRa
Cl_B_Yield_Reduction = Cl_B_Yield_Reduction + 10000 * DIRRb
Cl_A_Default_Rate = Cl_A_Default_Rate - 1 * (Worksheets("Liability").Range("Class_A_SP_Default_Rate").Value = True)
Cl_B_Default_Rate = Cl_B_Default_Rate - 1 * (Worksheets("Liability").Range("Class_B_SP_Default_Rate").Value = True)
'
' Weighted Average Life Calculations (Class A and Class B)



If Str(ALa) <> "Infinity" Then
WAL_A_Index = WAL_A_Index + 1
WAL_A = WAL_A + ALa
End If

If Str(ALb) <> "Infinity" Then
WAL_B_Index = WAL_B_Index + 1
WAL_B = WAL_B + ALb
End If

'
' Monitor if necessary

If Monitoring Then
Worksheets("Graphs").Cells(h + 1, 1).Value = h
Worksheets("Graphs").Cells(h + 1, 2).Value = IRR
End If
'
' Monitor convergence

Next h


' Compute the tranche-wise yield reductions and weighted average lives

Cl_A_Yield_Reduction = Cl_A_Yield_Reduction / Monte_Carlo_Runs
Cl_B_Yield_Reduction = Cl_B_Yield_Reduction / Monte_Carlo_Runs
Cl_A_Default_Rate = Cl_A_Default_Rate / Monte_Carlo_Runs
Cl_B_Default_Rate = Cl_B_Default_Rate / Monte_Carlo_Runs
WAL_A = WAL_A / WAL_A_Index
WAL_B = WAL_B / WAL_B_Index
'

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Non-Linear Rate Update

Yield_A = Yield_Curve(WAL_A, Cl_A_Yield_Reduction)
Yield_B = Yield_Curve(WAL_B, Cl_B_Yield_Reduction)
Delta_A = Yield_A - Cl_A_Rate
Delta_B = Yield_B - Cl_B_Rate
New_Class_A_Rate = Cl_A_Rate + Class_A_Relaxation_Factor * Delta_A
New_Class_B_Rate = Cl_B_Rate + Class_B_Relaxation_Factor * Delta_B


Convergence_Criterion = (Cl_A_Size / (Cl_A_Size + Cl_B_Size)) * Abs((New_Class_A_Rate - Cl_A_Rate) / Cl_A_Rate) + _
(1 - (Cl_A_Size / (Cl_A_Size + Cl_B_Size))) * Abs((New_Class_B_Rate - Cl_B_Rate) / Cl_B_Rate)


Cl_A_Rate = New_Class_A_Rate
Cl_B_Rate = New_Class_B_Rate
'
' Substitute the new updates into the spreadsheet for the next iteration


New_A = Cl_A_Rate
New_B = Cl_B_Rate
Worksheets("Liability").Range("Convergence_Status").Value = Non_Linear_Index


' Continue Until the non-Linear Convergence Criterion is Satisfied
Loop Until Application.Or(Convergence_Criterion < Epsilon, Non_Linear_Index > Non_Linear_Limit)


''''''''''''''''''''''''''''''''''''''''''''''''End of Non_Linear Loop''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'Write Out Results
'
If Non_Linear_Index > Non_Linear_Limit Then Worksheets("Liability").Range("Convergence_Status").Cells(2, 1).Value = "Diverged" Else _
Worksheets("Liability").Range("Convergence_Status").Cells(2, 1).Value = "Converged"

Moodys_Cl_A_Rating = Application.VLookup(Cl_A_Yield_Reduction, Rating_Scale_Array_M, 2, True)
Moodys_Cl_B_Rating = Application.VLookup(Cl_B_Yield_Reduction, Rating_Scale_Array_M, 2, True)
SP_Cl_A_Rating = Application.VLookup(100 * Cl_A_Default_Rate, Rating_Scale_Array_S, 2, True)
SP_Cl_B_Rating = Application.VLookup(100 * Cl_B_Default_Rate, Rating_Scale_Array_S, 2, True)
'

'Moody's ratings
'
Worksheets("Liability").Cells(44, 2).Value = Moodys_Cl_A_Rating
Worksheets("Liability").Cells(45, 2).Value = Moodys_Cl_B_Rating
Worksheets("Liability").Cells(48, 2).Value = Cl_A_Yield_Reduction
Worksheets("Liability").Cells(49, 2).Value = Cl_B_Yield_Reduction
'
'S&P's ratings
'
Worksheets("Liability").Cells(46, 2).Value = SP_Cl_A_Rating
Worksheets("Liability").Cells(47, 2).Value = SP_Cl_B_Rating
Worksheets("Liability").Cells(50, 2).Value = Cl_A_Default_Rate
Worksheets("Liability").Cells(51, 2).Value = Cl_B_Default_Rate

Worksheets("Liability").Cells(7, 5).Value = h - 1
End_Time = Now()
Elapsed_Time = (End_Time - Start_Time)
Worksheets("Liability").Cells(9, 5).Value = 86400 * Elapsed_Time


Application.Calculation = xlCalculationAutomatic
Application.DisplayStatusBar = True
Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub






Sub Inflation()
        product = 1
        For j = LBound(InflationAnnFactor) To UBound(InflationAnnFactor)

                AnnualInflationRate = WorksheetFunction.Beta_Inv(Rnd(), Inf_alpha, Inf_beta, Inf_Lower, Inf_Upper)
                product = product * (1 + AnnualInflationRate)

                InflationAnnFactor(j) = product
                InflationMonFactor(j) = AnnualInflationRate
                
        Next j
        
End Sub




Sub OperatingExp()

BasicOpex = WorksheetFunction.Beta_Inv(Rnd, Opex_alpha, Opex_beta, Opex_Lower, Opex_Upper)


If TodayAge(i) + t \ 12 >= 15 Then
    AddOpex = (TodayAge(i) + (t \ 12) - 15) / 15 * 0.05
Else
    AddOpex = 0
End If

TotalOpex = BasicOpex + AddOpex

End Sub




Sub NewLease()
 'draw new lease term and uninflated lease rate
 Dim res As Integer
 Dim rd As Integer
 
                                        If Age(i) >= 20 Then
                                            
                                                res = WorksheetFunction.VLookup(Capacity(i), myTableArray201, 2, True) \ 100
                                                
                                                If res < 4 Then
                                                rd = Int((100 * (res + 1) - 100 * res + 1) * Rnd() + 100 * res)
                                                Else
                                                rd = Int((467 - 100 * res + 1) * Rnd() + 100 * res)
                                                End If
                                                
                                                CharterLength(i) = WorksheetFunction.VLookup(rd, myTableArray202, 7, True)
                                                UninflatedRate(i) = WorksheetFunction.VLookup(rd, myTableArray202, 2, True)
                                                 
                                        ElseIf Age(i) >= 14 Then
                                        
                                                res = WorksheetFunction.VLookup(Capacity(i), myTableArray141, 2, True) \ 100
                                                
                                                If res < 9 Then
                                                rd = Int((100 * (res + 1) - 100 * res + 1) * Rnd() + 100 * res)
                                                Else
                                                rd = Int((931 - 100 * res + 1) * Rnd() + 100 * res)
                                                End If

                                                CharterLength(i) = WorksheetFunction.VLookup(rd, myTableArray142, 7, True)
                                                UninflatedRate(i) = WorksheetFunction.VLookup(rd, myTableArray142, 2, True)
                                               
                                               
                                        ElseIf Age(i) >= 9 Then
                                        
                                                res = WorksheetFunction.VLookup(Capacity(i), myTableArray091, 2, True) \ 100
                                                
                                                If res < 19 Then
                                                rd = Int((100 * (res + 1) - 100 * res + 1) * Rnd() + 100 * res)
                                                Else
                                                rd = Int((1933 - 100 * res + 1) * Rnd() + 100 * res)
                                                End If
                                                
                                                CharterLength(i) = WorksheetFunction.VLookup(rd, myTableArray092, 7, True)
                                                UninflatedRate(i) = WorksheetFunction.VLookup(rd, myTableArray092, 2, True)
                                              
                                              
                                        Else
                                        
                                                res = WorksheetFunction.VLookup(Capacity(i), myTableArray01, 2, True) \ 100
                                                
                                                If res < 6 Then
                                                rd = Int((100 * (res + 1) - 100 * res + 1) * Rnd() + 100 * res)
                                                Else
                                                rd = Int((688 - 100 * res + 1) * Rnd() + 100 * res)
                                                End If

                                                CharterLength(i) = WorksheetFunction.VLookup(rd, myTableArray02, 7, True)
                                                UninflatedRate(i) = WorksheetFunction.VLookup(rd, myTableArray02, 2, True)
                                                
                                        End If
End Sub
 



Sub RevenueFL()

'calculate revenue
    If StartT(i) + CharterLength(i) - 1 - t > 0 Then
            'calculate future cash flows in a month
            Revenue(i) = 30 * Rate 'cash inflow in a month
            Cost(i) = Revenue(i) * TotalOpex 'cash outflow in a months, e.g. operating expenses
            'Amortization of every ship each month

            'output results
            Total_Revenue(t) = Total_Revenue(t) + Revenue(i)
            Total_Cost(t) = Total_Cost(t) + Cost(i)
            Total_Amort(t) = Total_Amort(t) + Amort(i)
            Total_ScrapIncome(t) = Total_ScrapIncome(t) + 0 'ScrapIncome
            Total_Sales(t) = Total_Sales(t) + 0 'Sales
            
    ElseIf StartT(i) + CharterLength(i) - 1 - t = 0 Then
                                                       
            'calculate future cash flows in the last month of the lease
            Revenue(i) = 30 * Rate
            Cost(i) = Revenue(i) * TotalOpex
            
                   
            'output results
            Total_Revenue(t) = Total_Revenue(t) + Revenue(i)
            Total_Cost(t) = Total_Cost(t) + Cost(i)
            Total_Amort(t) = Total_Amort(t) + Amort(i)
            Total_ScrapIncome(t) = Total_ScrapIncome(t) + 0 'ScrapIncome
            Total_Sales(t) = Total_Sales(t) + 0 'Sales
            
            Lease(i) = 0
            Lease_Count(i) = Lease_Count(i) + 1
            
            'releasing intervals

            ReleaseInterval = Round((B - a) * Rnd + a, 0)
            ReleaseInterval_Countdown(i) = ReleaseInterval
    End If


End Sub



    
    
Sub ScrapInflation()
'inflation rates of scrap value

    Public AnnualSInflationRate As Double
    Dim u As Integer
            
        For u = 1 To t

                AnnualSInflationRate = WorksheetFunction.Beta_Inv(Rnd(), ScrapInf_alpha, ScrapInf_beta, ScrapInf_Lower, ScrapInf_Upper)
                SInflationFactor(i) = SInflationFactor(i) * AnnualSInflationRate
                
        Next u
                
End Sub                                    





Sub Sales()

'calculate discounted value
Public cf As Double
Public dcf As Double
Dim ratio As Double
Dim e As Integer

Public DepValue As Double
Public SalesPrice As Double
Public InfSalesPrice As Double
Public FinalSalesPrice As Double


        dcf = 0
        FinalSalesPrice = 0
       
        If Lease(i) = 1 Then
        
               cf = 30 * (Rate - Opex - SurExp)
               For e = 1 To (UsefulLife - Age(i)) * 12
                     dcf = dcf + cf / ((1 + WAI) ^ e)
               Next e
            
        ElseIf Lease(i) = 0 And Lease_Count(i) = 0 Then
         
                cf = 30 * (Rate - Opex - SurExp)
                For e = 1 To (UsefulLife - Age(i)) * 12
                     dcf = dcf + cf / ((1 + WAI) ^ e)
                Next e
        
        ElseIf Lease(i) > 1 Then
        
                cf = 30 * Rate * (1 - TotalOpex)
                For e = 1 To (UsefulLife - Age(i)) * 12
                    dcf = dcf + cf / ((1 + WAI) ^ e)
                Next e
        
        ElseIf Lease(i) = 0 And Lease_Count(i) > 1 Then
        
                cf = 30 * Rate * (1 - TotalOpex)
                For e = 1 To (UsefulLife - Age(i)) * 12
                    dcf = dcf + cf / ((1 + WAI) ^ e)
                Next e

        End If
            
        Call ScrapInflation
        
        ScrapValue(i) = sValue(i) * SInflationFactor(i)
        
        ''''''value 1 : scrap value
        ScrapIncome(i) = ScrapValue(i) * 1000000 '$ in dollars

        ratio = dcf / ScrapIncome(i)
        '''''value 2 : depreciated value
        DepValue = Amort(i) * ((UsefulLife - TodayAge(i)) * 12 - t)
        

        
        If Age(i) >= 0 And Age(i) <= 8 Then
            SalesPrice = WorksheetFunction.Beta_Inv(Rnd(), Sales1_alpha, Sales1_beta, Sales1_Lower, Sales1_Upper)
        ElseIf Age(i) <= 13 Then
            SalesPrice = WorksheetFunction.Beta_Inv(Rnd(), Sales2_alpha, Sales2_beta, Sales2_Lower, Sales2_Upper)
        ElseIf Age(i) <= 19 Then
            SalesPrice = WorksheetFunction.Beta_Inv(Rnd(), Sales3_alpha, Sales3_beta, Sales3_Lower, Sales3_Upper)
        Else
            SalesPrice = WorksheetFunction.Beta_Inv(Rnd(), Sales4_alpha, Sales4_beta, Sales4_Lower, Sales4_Upper)
        End If
        '''''value 3 : sales value
        InfSalesPrice = SalesPrice * SInflationFactor(i) * Capacity(i)
        
        ''''''Final interim sale price
        FinalSalesPrice = WorksheetFunction.Max(DepValue * (1 + ratio), ScrapIncome(i), InfSalesPrice)
        
        
        'outout resaults
        'Worksheets("Asset").Cells(3 + t, 7).Value = Worksheets("Asset").Cells(3 + t, 7).Value + FinalSalesPrice
        Total_Sales(t) = Total_Sales(t) + FinalSalesPrice

End Sub




Sub one_monte()

'Liability Array Redefine
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReDim CurrentCollection(Term), BegMR(Term), BegRA(Term)
ReDim BalA(Term), BalB(Term), BalC(Term)
ReDim s(Term), SP(Term), SS(Term), IA(Term), IAP(Term), IAS(Term)
ReDim PA(Term), PAP(Term), PAS(Term), PB(Term), PBP(Term), PBS(Term), PC(Term), PCP(Term), PCS(Term)
ReDim IB(Term), IBP(Term), IBS(Term), IC(Term), ICP(Term), ICS(Term)
ReDim EndMR(Term), RaM(Term), RaR(Term), RS(Term), EndRA(Term), RrR(Term), CICP(Term)
ReDim Fa(Term), Fb(Term), AAmort(Term), BAmort(Term), Month(Term)
ReDim F0(Term), F1(Term), F2(Term), F3(Term), F4(Term), F5(Term)
ReDim F6(Term), F7(Term), F8(Term), F9(Term), F10(Term), F11(Term), F12(Term)

'Ship Array Redefine
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

ReDim IsSold(Ships), HasBeenScrap(Ships), ScrapIncome(Ships)
ReDim Lease(Ships), Lease_Count(Ships), ReleaseInterval_Countdown(Ships)
ReDim ExecutedOnce(Ships), TodayAge(Ships), Capacity(Ships)
ReDim ScrapValue(Ships), Age(Ships), CharterLength(Ships), StartT(Ships)
ReDim UninflatedRate(Ships), Revenue(Ships), Cost(Ships), sValue(Ships)
ReDim Total_Revenue(Term), Total_Cost(Term), Total_Amort(Term), Total_ScrapIncome(Term), Total_Sales(Term)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Worksheets("Asset").Range("D4:AG363").ClearContents
'generate assets
Call Inflation  'generate the array of inflation rates

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Initialize Asset variables
For i = 1 To Ships

        Lease(i) = 1
        ReleaseInterval_Countdown(i) = 0
        ScrapIncome(i) = 0
        SInflationFactor(i) = 1

        IsSold(i) = False
        ExecutedOnce(i) = True
        HasBeenScrap(i) = False
            
        TodayAge(i) = TodayAge_initial(i) 'the age of the ship today
        Capacity(i) = Capacity_initial(i) 'in TEU nominal
        ScrapValue(i) = ScrapValue_initial(i) '$ in millions
        sValue(i) = sValue_initial(i)

Next i


'calcuate F0 at time0
F0(0) = 0 + MRperShip * Ships + RAratio * V  'current collectoin =0 before the first month
Fa(0) = -Cl_A_Bal
Fb(0) = -Cl_B_Bal

Cl_A_Rate = New_A
Cl_B_Rate = New_B

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'----loop----

'Time Loop
For ct = 1 To Term
t = ct


'Asset Part
'Ship Iteration

For i = 1 To Ships

        Age(i) = TodayAge(i) + (t \ 12) 'the age of each ship every month
        
                Select Case Age(i)
        
                        Case Is <= UsefulLife
                        
                        ''''''''''''''''Sales''''''''''''''''''
                            If IsSold(i) = False Then
                                
                                        ProbofSold = Rnd()

                                        '''''chosen for sale
                                        If (t = 24 Or ((t - 24) Mod 48 = 0)) And ProbofSold <= Sale_prob Then
                                                   
                                                   Call Sales
                                                   
                                                   Lease(i) = -1
                                                   'Worksheets("Asset").Cells(3 + t, 8 + i).Value = Lease(i)
                                                   IsSold(i) = True
                                                   HasBeenScrap(i) = True
                                                   
                                                   
                                        '''''''''not chosen for sale
                                        Else
                                                   'Worksheets("Asset").Cells(3 + t, 8 + i).Value = Lease(i)
                                                   If ReleaseInterval_Countdown(i) = 0 Then
                               
                                                   
                                                   ''''''''''''''''''''''''''''''Current Lease'''''''''''''''''''''''''''''''''
                                                           If Lease(i) = 1 Then
                        
                                                                       ExpirationDate = ExpirationDate_initial(i)
                                                                       CharterLength(i) = CharterLength_initial(i)  'Length of Lease (months)
                                                                       
                                                                       Rate = Rate_initial(i) '$/Day
                                                                       Opex = Opex_initial(i) '$/Day
                                                                       SurExp = SurExp_initial(i) 'Survey Expense $/Day
                                                                       
                                                                       If CharterLength(i) - t > 0 Then
                                                                               'calculate future cash flows in a month
                                                                               Revenue(i) = 30 * Rate 'cash inflow in a month
                                                                               Cost(i) = 30 * (Opex + SurExp) 'cash outflow in a months, e.g. operating expenses
                                                         
                                                                               Total_Revenue(t) = Total_Revenue(t) + Revenue(i)
                                                                               Total_Cost(t) = Total_Cost(t) + Cost(i)
                                                                               Total_Amort(t) = Total_Amort(t) + Amort(i)
                                                                               Total_ScrapIncome(t) = Total_ScrapIncome(t) + 0
                                                                               Total_Sales(t) = Total_Sales(t) + 0
                                                                               
                                                                               'Worksheets("Asset").Cells(3 + t, 4).Value = Total_Revenue(t)
                                                                               'Worksheets("Asset").Cells(3 + t, 5).Value = Total_Cost(t)
                                                                               'Worksheets("Asset").Cells(3 + t, 8).Value = Total_Amort(t)
                                                                               'Worksheets("Asset").Cells(3 + t, 6).Value = Total_ScrapIncome(t)
                                                                               'Worksheets("Asset").Cells(3 + t, 7).Value = Total_Sales(t)
                                                                               
                                                                       ElseIf CharterLength(i) - t = 0 Then
                                                                       
                                                                               'calculate future cash flows in the last month of the lease
                                                                               Revenue(i) = 30 * Rate
                                                                               Cost(i) = 30 * (Opex + SurExp)
                                                                               
                                                                               Total_Revenue(t) = Total_Revenue(t) + Revenue(i)
                                                                               Total_Cost(t) = Total_Cost(t) + Cost(i)
                                                                               Total_Amort(t) = Total_Amort(t) + Amort(i)
                                                                               Total_ScrapIncome(t) = Total_ScrapIncome(t) + 0
                                                                               Total_Sales(t) = Total_Sales(t) + 0
                                                                               
                                                                               'Worksheets("Asset").Cells(3 + t, 4).Value = Total_Revenue(t)
                                                                               'Worksheets("Asset").Cells(3 + t, 5).Value = Total_Cost(t)
                                                                               'Worksheets("Asset").Cells(3 + t, 8).Value = Total_Amort(t)
                                                                               'Worksheets("Asset").Cells(3 + t, 6).Value = Total_ScrapIncome(t)
                                                                               'Worksheets("Asset").Cells(3 + t, 7).Value = Total_Sales(t)
                                                                               
                                                                               Lease(i) = 0
                                                                               Lease_Count(i) = 1
                                                                               
                                                                               'releasing intervals
                                                                            
                                                                               ReleaseInterval = Round((B - a) * Rnd + a, 0)
                                                                               ReleaseInterval_Countdown(i) = ReleaseInterval
                                                                                                                                            
                                                                       End If
                                                               
                                                       ''''''''''''''''''''''''''''''''''future leases''''''''''''''''''''''''''''''''''
                                                       ElseIf Lease(i) > 1 Then
                                                       
                                                               If ExecutedOnce(i) = True Then
                                                               
                                                                          StartDate = DateAdd("m", t, Date)
                                                                          StartT(i) = t
                        
                                                                          'draw new lease term and uninflated lease rate
                                                                          Call NewLease
                                                                          
                                                                          'pick inflation rate
                                                                          Y = t \ 12
                                                                          M = t - Y * 12
                                                                          
                                                                          If Y > 0 Then
                                                                                InflationRate = InflationAnnFactor(Y) * (1 + (M / 12) * InflationMonFactor(Y + 1))
                                                                          Else
                                                                                InflationRate = 1 + (M / 12) * InflationMonFactor(Y + 1)
                                                                          End If
                                                                          
                                                                          'calculate inflated future lease rate
                                                                          Rate = UninflatedRate(i) * InflationRate
                                                               
                                                                          'operating expenses (in ratio)
                                                                          Call OperatingExp
                                                                          
                                                                          ExecutedOnce(i) = False

                                                                          Call RevenueFL

                                                               Else 'ExecutedOnce=False
                                                                          Call RevenueFL
                                                            

                                                                End If
                                                                
                                                        End If
                                   
                                                   Else 'ReleaseInterval_Countdown(i) NOT =0

                                                           Revenue(i) = 0
                                                           Cost(i) = 0
                                                           
                                                           Total_Revenue(t) = Total_Revenue(t) + Revenue(i)
                                                           Total_Cost(t) = Total_Cost(t) + Cost(i)
                                                           Total_Amort(t) = Total_Amort(t) + Amort(i)
                                                           Total_ScrapIncome(t) = Total_ScrapIncome(t) + 0
                                                           Total_Sales(t) = Total_Sales(t) + 0
                                                           
                                                           'Worksheets("Asset").Cells(3 + t, 4).Value = Total_Revenue(t)
                                                           'Worksheets("Asset").Cells(3 + t, 5).Value = Total_Cost(t)
                                                           'Worksheets("Asset").Cells(3 + t, 8).Value = Total_Amort(t)
                                                           'Worksheets("Asset").Cells(3 + t, 6).Value = Total_ScrapIncome(t)
                                                           'Worksheets("Asset").Cells(3 + t, 7).Value = Total_Sales(t)
                                                                               
                                                           ReleaseInterval_Countdown(i) = ReleaseInterval_Countdown(i) - 1
                                                               
                                                           If ReleaseInterval_Countdown(i) = 0 Then
                                                               Lease(i) = Lease_Count(i) + 1
                                                           Else
                                                               Lease(i) = 0
                                                           End If
                                                           
                                                           
                                                           ExecutedOnce(i) = True
                                                           
                                                   End If
                                              
                                              
                                        End If
                                    
                                    
                            ElseIf IsSold(i) = True Then
                                        'Worksheets("Asset").Cells(3 + t, 8 + i).Value = Lease(i)
                                        Total_Revenue(t) = Total_Revenue(t)
                                        Total_Cost(t) = Total_Cost(t)
                                        Total_Amort(t) = Total_Amort(t)
                                        Total_ScrapIncome(t) = Total_ScrapIncome(t)
                                        Total_Sales(t) = Total_Sales(t)
                                        
                                        'Worksheets("Asset").Cells(3 + t, 4).Value = Total_Revenue(t)
                                        'Worksheets("Asset").Cells(3 + t, 5).Value = Total_Cost(t)
                                        'Worksheets("Asset").Cells(3 + t, 8).Value = Total_Amort(t)
                                        'Worksheets("Asset").Cells(3 + t, 6).Value = Total_ScrapIncome(t)
                                        'Worksheets("Asset").Cells(3 + t, 7).Value = Total_Sales(t)
                                        
                                        HasBeenScrap(i) = True
                                        IsSold(i) = True
                                        
                            End If
                                
                       '''''''''''''''''''''''''''Ship age above 30'''''''''''''''''''''''
                        Case Else
                    
                                If HasBeenScrap(i) = False Then
                                
                                    Call ScrapInflation
                                    
                                    ScrapValue(i) = sValue(i) * SInflationFactor(i)
                                    ScrapIncome(i) = ScrapValue(i) * 1000000 '$ in dollars
                                    Total_ScrapIncome(t) = Total_ScrapIncome(t) + ScrapIncome(i)
                                    HasBeenScrap(i) = True
                                    

                                Else

                                     Total_ScrapIncome(t) = Total_ScrapIncome(t)
                                End If
                                
                        
                        End Select

Next i

Month(t) = t

'Calculate Current Collection each month
CurrentCollection(t) = Total_Revenue(t) + Total_ScrapIncome(t) + Total_Sales(t) - Total_Cost(t)

'calculate beginning maintenance reserve and interest reserve accounts after time 0
If t = 1 Then
    BegMR(t) = initial_maintenance_reserve * (1 + IR / 12)
    BegRA(t) = initial_interest_reserve * (1 + IR / 12)
Else
    BegMR(t) = EndMR(t - 1) * (1 + IR / 12)
    BegRA(t) = EndRA(t - 1) * (1 + IR / 12)
End If

F0(t) = CurrentCollection(t) + BegMR(t) + BegRA(t)


'''''''''''''''''''''''''''''Part1:Servicing Fee''''''''''''''''''''''''''''''
'number of ships out of lease
Dim itr, count As Integer
count = 0
For itr = 1 To Ships
If IsSold(itr) Or HasBeenScrap(itr) Then count = count + 1
Next itr

'Worksheets("Asset").Cells(3 + t, 37).Value = count
N = Ships - count

'Service fee due
If t = 1 Then
    s(t) = MRperShip / 12 * N
Else
    s(t) = MRperShip / 12 * N + SS(t - 1)
End If
    
SP(t) = WorksheetFunction.Min(F0, s(t)) 'Service fee paid
SS(t) = s(t) - SP(t) 'Service fee shorfall
F1(t) = F0(t) - SP(t)

'''''''''Draw Maintenance Fees
RdM = WorksheetFunction.Max(0, (BegMR(t) - F1(t))) 'Maintenance reserve draw
F2(t) = WorksheetFunction.Max(0, (F1(t) - BegMR(t)))


'''''''''''''''''''''''''''''Part2:Class A Interest''''''''''''''''''''''''''''''
'Class A interest due
If t = 1 Then
    IA(t) = (Cl_A_Rate / 12) * Cl_A_Bal
Else
    IA(t) = (Cl_A_Rate / 12) * BalA(t - 1) + IAS(t - 1) * (1 + Cl_A_Rate / 12)
End If

IAP(t) = WorksheetFunction.Min(F2, IA(t)) 'class A interest paid
IAS(t) = IA(t) - IAP(t) 'Class A interest shortfall
F3(t) = F2(t) - IAP(t)


'''''''''''''''''''''''''''''Part3:Class B Interest''''''''''''''''''''''''''''''
'Class B interest due
If t = 1 Then
    IB(t) = (Cl_B_Rate / 12) * Cl_B_Bal
Else
    IB(t) = (Cl_B_Rate / 12) * BalB(t - 1) + IBS(t - 1) * (1 + Cl_B_Rate / 12)
End If

IBP(t) = WorksheetFunction.Min(F3, IB(t)) 'class B interest paid
IBS(t) = IB(t) - IBP(t) 'Class B interest shortfall
F4(t) = F3(t) - IBP(t)

'''''''''Draw Interest Reserve
RdR = WorksheetFunction.Max(0, (BegRA(t) - F4(t))) 'Interest reserve draw
F5(t) = WorksheetFunction.Max(0, (F4(t) - BegRA(t)))


'''''''''''''''''''''''''''''Part4:Top Up Maintenance Reserve Account''''''''''''''''''''''''''''''
RrM = N * MRperShip
RaM(t) = WorksheetFunction.Min(F5(t), (RrM - BegMR(t) + RdM))

EndMR(t) = BegMR(t) - RdM + RaM(t)
F6(t) = F5(t) - RaM(t)


'''''''''''''''''''''''''''''Part5:Class C Interest''''''''''''''''''''''''''''''
'Class C interest due
If t = 1 Then
    IC(t) = (Cl_C_Rate / 12) * Cl_C_Bal
Else
    IC(t) = (Cl_C_Rate / 12) * BalC(t - 1) + ICS(t - 1) * (1 + Cl_C_Rate / 12)
End If

ICP(t) = WorksheetFunction.Min(F6, IC(t)) 'class C interest paid
ICS(t) = IC(t) - ICP(t) 'Class C interest shortfall
F7(t) = F6(t) - ICP(t)


'''''''''''''''''''''''''''''Part6:Top Up Interest Reserve Account''''''''''''''''''''''''''''''
If t = 1 Then
    RrR(t) = (WAImonth / 12) * (Cl_A_Bal * Cl_A_Rate + Cl_B_Bal * Cl_B_Rate)
Else
    RrR(t) = (WAImonth / 12) * (BalA(t - 1) * Cl_A_Rate + BalB(t - 1) * Cl_B_Rate)
End If

RaR(t) = WorksheetFunction.Min(F7, (RrR(t) - BegRA(t) + RdR))
EndRA(t) = BegRA(t) - RdR + RaR(t)
F8(t) = F7(t) - RaR(t)


'''''''''''''''''''''''''''''Part7:Class A Principal''''''''''''''''''''''''''''''
If t = 1 Then
    PA(t) = Cl_A_Bal - TACa(t)
Else
    PA(t) = BalA(t - 1) - TACa(t)
End If

PAP(t) = WorksheetFunction.Min(F8(t), PA(t))
PAS(t) = PA(t) - PAP(t)
F9(t) = F8(t) - PAP(t)

If t = 1 Then
    BalA(t) = Cl_A_Bal - PAP(t)
    AAmort(t) = Cl_A_Bal - BalA(t)
Else
    BalA(t) = BalA(t - 1) - PAP(t)
    AAmort(t) = BalA(t - 1) - BalA(t)
End If


'''''''''''''''''''''''''''''Part8:Class B Principal''''''''''''''''''''''''''''''
If t = 1 Then
    PB(t) = Cl_B_Bal - TACb(t)
Else
    PB(t) = BalB(t - 1) - TACb(t)
End If

PBP(t) = WorksheetFunction.Min(F9(t), PB(t))
PBS(t) = PB(t) - PBP(t)
F10(t) = F9(t) - PBP(t)

If t = 1 Then
    BalB(t) = Cl_B_Bal - PBP(t)
    BAmort(t) = Cl_B_Bal - BalB(t)
Else
    BalB(t) = BalB(t - 1) - PBP(t)
    BAmort(t) = BalB(t - 1) - BalB(t)
End If


'''''''''''''''''''''''''''''Part9:Class C ''''''''''''''''''''''''''''''
If t = 1 Then
    PC(t) = Cl_C_Bal
Else
    PC(t) = BalC(t - 1)
End If

    PCP(t) = WorksheetFunction.Min(F10(t), PC(t))
    PCS(t) = PC(t) - PCP(t)
    F11(t) = F10(t) - PCP(t)

If t = 1 Then
    BalC(t) = Cl_C_Bal - PCP(t) 'Calculate class principals after payment
Else
    BalC(t) = BalC(t - 1) - PCP(t)
End If



'''''''''''''''''''''''''''''Part10:Residual''''''''''''''''''''''''''''''
RS(t) = F11(t)

If t = 1 Then
    CICP(t) = RS(t)
Else
    CICP(t) = CICP(t - 1) + RS(t)
End If

F12(t) = F11(t) - RS(t)

'calculate total payments
Fa(t) = IAP(t) + PAP(t)
Fb(t) = IBP(t) + PBP(t)


''''''''''''''''''''''''''''''''''''END OF WATERFALL'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Dim its As Integer
'its = Worksheets("Sheet1").Cells(1, 1).Value

'    Worksheets("Sheet1").Cells(3 + t, 1).Value = t
'    Worksheets("Sheet1").Cells(3 + t, 2).Value = Age(its)
'    Worksheets("Sheet1").Cells(3 + t, 3).Value = UninflatedRate(its)
'    Worksheets("Sheet1").Cells(3 + t, 4).Value = Revenue(its)
'    Worksheets("Sheet1").Cells(3 + t, 5).Value = Cost(its)
'    Worksheets("Sheet1").Cells(3 + t, 6).Value = Amort(its)
'    Worksheets("Sheet1").Cells(3 + t, 7).Value = ScrapValue(its)
'    Worksheets("Sheet1").Cells(3 + t, 8).Value = ScrapIncome(its)
'    Worksheets("Sheet1").Cells(3 + t, 9).Value = Lease(its)
'    Worksheets("Sheet1").Cells(3 + t, 10).Value = CharterLength(its)
'    Worksheets("Sheet1").Cells(3 + t, 11).Value = ReleaseInterval_Countdown(its)
      

Next ct



'calculate IRR for class a and b
IRRa = WorksheetFunction.IRR(Fa(), Cl_A_Rate / 12)
IRRb = WorksheetFunction.IRR(Fb(), Cl_B_Rate / 12)

'calculate Delta irr
DIRRa = WorksheetFunction.Max(0, Cl_A_Rate - WorksheetFunction.Max(-1, 12 * IRRa))
DIRRb = WorksheetFunction.Max(0, Cl_B_Rate - WorksheetFunction.Max(-1, 12 * IRRb))


'calculate average life
'If BalA(UBound(BalA)) < 1 Then
    ALa = WorksheetFunction.SumProduct(Month(), AAmort()) / Cl_A_Bal
'Else
    'ALa = "Infinity"
'End If

'If BalB(UBound(BalB)) < 1 Then
    ALb = WorksheetFunction.SumProduct(Month(), BAmort()) / Cl_B_Bal
'Else
    'ALb = "Infinity"
'End If

F11(0) = -Cl_C_Bal
IRR = 12 * WorksheetFunction.IRR(F11(), -0.1)


End Sub




