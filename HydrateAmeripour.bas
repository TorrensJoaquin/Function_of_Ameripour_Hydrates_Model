Option Explicit
'Molecular Weights
Const MWH2S As Double = 34.08
Const MWCO2 As Double = 44.01
Const MWN2 As Double = 28.01
Const MWC1 As Double = 16.04
Const MWC2 As Double = 30.07
Const MWC3 As Double = 44.1
Const MWiC4 As Double = 58.12
Const MWnC4 As Double = 58.12
Const MWiC5 As Double = 72.15
Const MWnC5 As Double = 72.15
Const MWnC6 As Double = 86.18
Const MWnC7 As Double = 100.26
Const MWnC8 As Double = 114.23
Const MWC2H4 As Double = 28.05
Const MWC3H6 As Double = 42.08
Const MWNaCl As Double = 58.448
Const MWKCl As Double = 74.551
Const MWCaCl2 As Double = 110.986
Const MWCH3OH As Double = 32.043
Const MWEG As Double = 62.07
Const MWTEG As Double = 150.2
Const MWGL As Double = 92
'Tc of Gas Component in Rankine
Const TcH2S As Double = 672.45
Const TcCO2 As Double = 547.91
Const TcN2 As Double = 227.49
Const TcC1 As Double = 343.33
Const TcC2 As Double = 549.92
Const TcC3 As Double = 666.06
Const TciC4 As Double = 734.46
Const TcnC4 As Double = 765.62
Const TciC5 As Double = 829.1
Const TcnC5 As Double = 845.8
Const TcnC6 As Double = 913.6
Const TcnC7 As Double = 972.7
Const TcnC8 As Double = 1024.22
Const TcC2H4 As Double = 508.58
Const TcC3H6 As Double = 656.9
'pc of Gas Component in psia
Const pcH2S As Double = 1300
Const pcCO2 As Double = 1071
Const pcN2 As Double = 493.1
Const pcC1 As Double = 666.4
Const pcC2 As Double = 706.5
Const pcC3 As Double = 616
Const pciC4 As Double = 527.9
Const pcnC4 As Double = 550.6
Const pciC5 As Double = 490.4
Const pcnC5 As Double = 488.6
Const pcnC6 As Double = 436.9
Const pcnC7 As Double = 396.8
Const pcnC8 As Double = 360.7
Const pcC2H4 As Double = 729.8
Const pcC3H6 As Double = 669
'J Constants
Const j0 As Double = 0.052073
Const j1 As Double = 1.016
Const j2 As Double = 0.86961
Const j3 As Double = 0.72646
Const j4 As Double = 0.85101
'K Constants
Const k0 As Double = -0.39741
Const k1 As Double = 1.0503
Const k2 As Double = 0.96592
Const K3 As Double = 0.78569
Const k4 As Double = 0.98211
'ai: Constants to calculate hydrate pseudoreduced pressure
Const a0 As Double = -2.924729
Const a1 As Double = 7.069408
Const a2 As Double = -0.671674
Const a3 As Double = 2.158912
Const a4 As Double = -0.014446
Const a5 As Double = 3.367516
Const a6 As Double = -0.168816
Const a7 As Double = 13.213962
Const a8 As Double = 2.365031
Const a9 As Double = -0.025796
Const a10 As Double = 2.461102
Const a11 As Double = -7.100059
Const a12 As Double = 1.820312
Const a13 As Double = 7.517561
Const a14 As Double = -0.018793
Const a15 As Double = 0.019029
Const a16 As Double = -0.005307
Const a17 As Double = -0.032564
'bi: Constants to calculate hydrate pseudoreduced presure
Const b0 As Double = 3.1113797464
Const b1 As Double = -0.06121811
Const b2 As Double = -0.034581592
Const b3 As Double = -0.022257841
Const b4 As Double = -0.161387206
Const b5 As Double = 0.0004644864
Const b6 As Double = 0.0060870675
Const b7 As Double = -0.00049726
Const b8 As Double = 0.0001682281
Const b9 As Double = -0.193610096
Const b10 As Double = 0.0001963793
Const b11 As Double = 0.1324677497
Const b12 As Double = -0.078512137
Const b13 As Double = 0.009232805
Const b14 As Double = -0.000232276
Const b15 As Double = 0.8054836679
Const b16 As Double = 0.0063403148
'WVP: Constants to calculate water vapor pressure
Const ap As Double = 29.86
Const bp As Double = -3152
Const Cp As Double = -7.3037
Const dp As Double = 0.000000002424
Const ep As Double = 0.000001809
'WLV: Constants to calculate water liquid viscosity
Const av As Double = -10.2158
Const bv As Double = 1792
Const Cv As Double = 0.0177
Const dv As Double = -0.0000126
Function HydrateAmeripourS(TempInF As Double, PressInPSI As Double, LabelsOfComponents As Range, ValuesOfComponents_MolPercentage As Range) As Variant
Application.Volatile
Dim i As Integer
Dim x As Byte
Dim numberOfComponents As Integer
Dim CheckForItemsNotRecognized() As Variant
Dim Labels() As Variant
'
Dim AUX As Variant
'Input data
Dim T As Double, P As Double, H2S As Double, CO2 As Double, n2 As Double, c1 As Double
Dim c2 As Double, C3 As Double, iC4 As Double, nC4 As Double, iC5 As Double, nC5 As Double
Dim nC6 As Double, nC7 As Double, nC8 As Double, C2H4 As Double, C3H6 As Double
Dim NaCl As Double, KCl As Double, CaCl2 As Double, CH3OH As Double, EG As Double
Dim TEG As Double, GL As Double
Dim j As Double, k As Double, Tpc As Double, Ppc As Double, Tpr As Double, Ppr As Double
Dim GSG As Double, lnGSG As Double, lnTpr As Double, Thr As Double, lnP As Double, Th As Double
Dim Ta As Double, Tb As Double, WVP As Double, WLV As Double
Dim Phr As Double, Ph As Double
'Check if we are on range
If TempInF > 90 Then HydrateAmeripourS = "": Exit Function
If PressInPSI > 12000 Then HydrateAmeripourS = "": Exit Function
'
numberOfComponents = LabelsOfComponents.Columns.Count
'This part of the code will check if we are working with columns or rows
If numberOfComponents = 1 Then
    numberOfComponents = LabelsOfComponents.Rows.Count
    ReDim Labels(1 To numberOfComponents)
    Labels = WorksheetFunction.Transpose(LabelsOfComponents.Value2)
    'I want to display an error if the size of the labels is different to the size of the values
    If numberOfComponents <> ValuesOfComponents_MolPercentage.Rows.Count Then HydrateAmeripourS = "Problem in data Selection": Exit Function
Else
    ReDim Labels(1 To numberOfComponents)
    For x = 1 To numberOfComponents
        Labels(x) = LabelsOfComponents(x)
    Next x
    'I want to display an error if the size of the labels is different to the size of the values
    If numberOfComponents <> ValuesOfComponents_MolPercentage.Columns.Count Then HydrateAmeripourS = "Problem in data Selection": Exit Function
End If
ReDim CheckForItemsNotRecognized(1 To numberOfComponents)
'Convert the ranges in variables.
For i = 1 To numberOfComponents
    ReDim AUX(1 To 3)
    AUX(1) = "H2S"
    AUX(2) = "SH2"
    AUX(3) = "HYDROGEN SULFIDE"
    If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then H2S = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
    ReDim AUX(1 To 3)
    AUX(1) = "CO2"
    AUX(2) = "CARBON DIOXIDE"
    AUX(3) = "CARBONIC ACID"
    If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then CO2 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
    ReDim AUX(1 To 3)
    AUX(1) = "N2"
    AUX(2) = "NITROGEN"
    AUX(3) = "N"
    If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then n2 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
    ReDim AUX(1 To 3)
    AUX(1) = "C1"
    AUX(2) = "CH4"
    AUX(3) = "METHANE"
    If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then c1 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
    ReDim AUX(1 To 3)
    AUX(1) = "C2"
    AUX(2) = "C2H6"
    AUX(3) = "ETHANE"
    If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then c2 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
    ReDim AUX(1 To 3)
    AUX(1) = "C3"
    AUX(2) = "C3H8"
    AUX(3) = "PROPANE"
    If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then C3 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
    ReDim AUX(1 To 3)
    AUX(1) = "IC4"
    AUX(2) = "IC4H10"
    AUX(3) = "ISOBUTANE"
    If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then iC4 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
    ReDim AUX(1 To 6)
    AUX(1) = "NC4"
    AUX(2) = "C4"
    AUX(3) = "NC4H10"
    AUX(4) = "NORMALBUTANE"
    AUX(5) = "NORMAL BUTANE"
    AUX(6) = "BUTANE"
    If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then nC4 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
    ReDim AUX(1 To 3)
    AUX(1) = "IC5"
    AUX(2) = "IC5H12"
    AUX(3) = "ISOPENTANE"
    If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then iC5 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
    ReDim AUX(1 To 6)
    AUX(1) = "NC5"
    AUX(2) = "C5"
    AUX(3) = "NC5H12"
    AUX(4) = "NORMALPENTANE"
    AUX(5) = "NORMAL PENTANE"
    AUX(6) = "PENTANE"
    If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then nC5 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
    ReDim AUX(1 To 4)
    AUX(1) = "NC6"
    AUX(2) = "C6"
    AUX(3) = "C6H14"
    AUX(4) = "HEXANE"
    If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then nC6 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
    ReDim AUX(1 To 4)
    AUX(1) = "NC7"
    AUX(2) = "C7"
    AUX(3) = "C7H16"
    AUX(4) = "HEPTANE"
    If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then nC7 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
    ReDim AUX(1 To 5)
    AUX(1) = "NC8"
    AUX(2) = "C8"
    AUX(3) = "C8H18"
    AUX(4) = "C8H18+"
    AUX(5) = "OCTANE"
    If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then nC8 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
    ReDim AUX(1 To 2)
    AUX(1) = "C2H4"
    AUX(2) = "ETHENE"
    If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then C2H4 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
    ReDim AUX(1 To 2)
    AUX(1) = "C3H6"
    AUX(2) = "PROPENE"
    If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then C3H6 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
    ReDim AUX(1 To 3)
    AUX(1) = "NACL"
    AUX(2) = "SALT"
    AUX(3) = "SODIUM CHLORIDE"
    If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then NaCl = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
    ReDim AUX(1 To 2)
    AUX(1) = "KCL"
    AUX(2) = "POTASSIUM CHLORIDE"
    If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then KCl = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
    ReDim AUX(1 To 3)
    AUX(1) = "CACL"
    AUX(2) = "CALCIUM CHLORIDE"
    AUX(2) = "CACL2"
    If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then CaCl2 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
    ReDim AUX(1 To 2)
    AUX(1) = "CH3OH"
    AUX(2) = "METHANOL"
    If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then CH3OH = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
    ReDim AUX(1 To 2)
    AUX(1) = "EG"
    AUX(2) = "ETHYLENE GLYCOL"
    If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then EG = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
    ReDim AUX(1 To 2)
    AUX(1) = "TEG"
    AUX(2) = "TRIETHYLENE GLYCOL"
    If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then TEG = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
    ReDim AUX(1 To 2)
    AUX(1) = "GL"
    AUX(2) = "GLYCOL"
    If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then GL = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
NextIteration:
    Erase AUX
Next i
'
'I want to display an error if a component is not recognized
For i = 1 To numberOfComponents
    If CheckForItemsNotRecognized(i) <> 1 Then
        HydrateAmeripourS = "ItemDontRecognized": Exit Function
    End If
Next i
'Normalization to 100%
Dim SumOfConcentrations As Double
SumOfConcentrations = (H2S + CO2 + n2 + c1 + c2 + C3 + iC4 + nC4 + iC5 + nC5 + nC6 + nC7 + nC8 + C2H4 + C3H6) / 100
H2S = H2S / SumOfConcentrations
CO2 = CO2 / SumOfConcentrations
n2 = n2 / SumOfConcentrations
c1 = c1 / SumOfConcentrations
c2 = c2 / SumOfConcentrations
C3 = C3 / SumOfConcentrations
iC4 = iC4 / SumOfConcentrations
nC4 = nC4 / SumOfConcentrations
iC5 = iC5 / SumOfConcentrations
nC5 = nC5 / SumOfConcentrations
nC6 = nC6 / SumOfConcentrations
nC7 = nC7 / SumOfConcentrations
nC8 = nC8 / SumOfConcentrations
C2H4 = C2H4 / SumOfConcentrations
C3H6 = C3H6 / SumOfConcentrations
'Check if we are on Range
If NaCl > 20 Then HydrateAmeripourS = "": Exit Function
If KCl > 20 Then HydrateAmeripourS = "": Exit Function
If CaCl2 > 20 Then HydrateAmeripourS = "": Exit Function
If CH3OH > 20 Then HydrateAmeripourS = "": Exit Function
If EG > 40 Then HydrateAmeripourS = "": Exit Function
If TEG > 40 Then HydrateAmeripourS = "": Exit Function
If GL > 40 Then HydrateAmeripourS = "": Exit Function
'
If TempInF = 0 And PressInPSI <> 0 Then
    For x = 1 To 20
        P = PressInPSI * x / 20
'Mixing Rules for pseudocritical temperature and presuure (gas component is in fraction)
        j = j0 + (j1 * (H2S / 100)) * (TcH2S / pcH2S) + (j2 * (CO2 / 100)) * (TcCO2 / pcCO2) + (j3 * (n2 / 100)) * (TcN2 / pcN2) + (j4 * (c1 / 100)) * (TcC1 / pcC1) + (j4 * (c2 / 100)) * (TcC2 / pcC2) + (j4 * (C3 / 100)) * (TcC3 / pcC3) + (j4 * (iC4 / 100)) * (TciC4 / pciC4) + (j4 * (nC4 / 100)) * (TcnC4 / pcnC4) + (j4 * (iC5 / 100)) * (TciC5 / pciC5) + (j4 * (nC5 / 100)) * (TcnC5 / pcnC5) + (j4 * (nC6 / 100)) * (TcnC6 / pcnC6) + (j4 * (nC7 / 100)) * (TcnC7 / pcnC7) + (j4 * (nC8 / 100)) * (TcnC8 / pcnC8) + (j4 * (C2H4 / 100)) * (TcC2H4 / pcC2H4) + (j4 * (C3H6 / 100)) * (TcC3H6 / pcC3H6)
        k = k0 + (k1 * (H2S / 100)) * (TcH2S / (pcH2S) ^ 0.5) + (k2 * (CO2 / 100)) * (TcCO2 / (pcCO2) ^ 0.5) + (K3 * (n2 / 100)) * (TcN2 / (pcN2) ^ 0.5) + (k4 * (c1 / 100)) * (TcC1 / (pcC1) ^ 0.5) + (k4 * (c2 / 100)) * (TcC2 / (pcC2) ^ 0.5) + (k4 * (C3 / 100)) * (TcC3 / (pcC3) ^ 0.5) + (k4 * (iC4 / 100)) * (TciC4 / (pciC4) ^ 0.5) + (k4 * (nC4 / 100)) * (TcnC4 / (pcnC4) ^ 0.5) + (k4 * (iC5 / 100)) * (TciC5 / (pciC5) ^ 0.5) + (k4 * (nC5 / 100)) * (TcnC5 / (pcnC5) ^ 0.5) + (k4 * (nC6 / 100)) * (TcnC6 / (pcnC6) ^ 0.5) + (k4 * (nC7 / 100)) * (TcnC7 / (pcnC7) ^ 0.5) + (k4 * (nC8 / 100)) * (TcnC8 / (pcnC8) ^ 0.5) + (k4 * (C2H4 / 100)) * (TcC2H4 / (pcC2H4) ^ 0.5) + (k4 * (C3H6 / 100)) * (TcC3H6 / (pcC3H6) ^ 0.5)
'Tpc: Pseudocritical Temperature
        Tpc = k * k / j
'Ppc: Pseudocritical Pressure
        Ppc = Tpc / j
'Ppr: Pseudoreduced Pressure
        Ppr = (P + 14.7) / Ppc
'GSG: Gas Specific Gravity Calculation
        GSG = ((H2S * MWH2S + CO2 * MWCO2 + n2 * MWN2 + c1 * MWC1 + c2 * MWC2 + C3 * MWC3 + iC4 * MWiC4 + nC4 * MWnC4 + (iC5 * MWiC5) + nC5 * MWnC5 + nC6 * MWnC6 + nC7 * MWnC7 + nC8 * MWnC8 + C2H4 * MWC2H4 + C3H6 * MWC3H6) / 100) / 29
        lnGSG = Log(GSG)
'
        lnTpr = (b0 + b1 * (Log(P) ^ 2) + b2 * ((NaCl / MWNaCl + KCl / MWKCl + CaCl2 / MWCaCl2) / (GSG ^ 2)) + b3 * ((CH3OH / MWCH3OH + EG / MWEG + TEG / MWTEG + GL / MWTEG) / (GSG ^ 2)) + b4 * (GSG ^ 2) + b5 * ((100 - (NaCl + KCl + CaCl2)) * (GSG ^ 3)) + b6 * (H2S + CO2 + n2) + b7 * ((CH3OH / MWCH3OH + EG / MWEG + TEG / MWTEG + GL / MWTEG) * (H2S + CO2 + n2)) + b8 * (C3 + iC4) / (GSG ^ 6) + b9 * ((Log(GSG)) * (Log(P))) + b10 * ((Log(GSG)) * (Log(P) ^ 4)) + b11 * ((Log(P)) / GSG) + b12 * (Log(P) ^ 2 / GSG) + b13 * (Log(P) ^ 3 / GSG) + b14 * (Log(P) ^ 4 / GSG) + b15 * (Log(Ppr)) + b16 * (Log(Ppr) ^ 2))
'Thr: Hydrate Pseudoreduced Temperature
        Thr = Exp(lnTpr)
'Th: Hydrate-Formation Temperature
        Th = (Thr * Tpc) - 460
        If HydrateAmeripourS > Th Then HydrateAmeripourS = "": Exit Function
        HydrateAmeripourS = Th
    Next x
    If HydrateAmeripourS > 90 Then HydrateAmeripourS = "": Exit Function
    Exit Function
End If
If TempInF <> 0 And PressInPSI = 0 Then
    For x = 1 To 20
        T = TempInF * x / 20
'Mixing Rules for pseudocritical temperature and presuure (gas component is in fraction)
        j = j0 + (j1 * (H2S / 100)) * (TcH2S / pcH2S) + (j2 * (CO2 / 100)) * (TcCO2 / pcCO2) + (j3 * (n2 / 100)) * (TcN2 / pcN2) + (j4 * (c1 / 100)) * (TcC1 / pcC1) + (j4 * (c2 / 100)) * (TcC2 / pcC2) + (j4 * (C3 / 100)) * (TcC3 / pcC3) + (j4 * (iC4 / 100)) * (TciC4 / pciC4) + (j4 * (nC4 / 100)) * (TcnC4 / pcnC4) + (j4 * (iC5 / 100)) * (TciC5 / pciC5) + (j4 * (nC5 / 100)) * (TcnC5 / pcnC5) + (j4 * (nC6 / 100)) * (TcnC6 / pcnC6) + (j4 * (nC7 / 100)) * (TcnC7 / pcnC7) + (j4 * (nC8 / 100)) * (TcnC8 / pcnC8) + (j4 * (C2H4 / 100)) * (TcC2H4 / pcC2H4) + (j4 * (C3H6 / 100)) * (TcC3H6 / pcC3H6)
        k = k0 + (k1 * (H2S / 100)) * (TcH2S / (pcH2S) ^ 0.5) + (k2 * (CO2 / 100)) * (TcCO2 / (pcCO2) ^ 0.5) + (K3 * (n2 / 100)) * (TcN2 / (pcN2) ^ 0.5) + (k4 * (c1 / 100)) * (TcC1 / (pcC1) ^ 0.5) + (k4 * (c2 / 100)) * (TcC2 / (pcC2) ^ 0.5) + (k4 * (C3 / 100)) * (TcC3 / (pcC3) ^ 0.5) + (k4 * (iC4 / 100)) * (TciC4 / (pciC4) ^ 0.5) + (k4 * (nC4 / 100)) * (TcnC4 / (pcnC4) ^ 0.5) + (k4 * (iC5 / 100)) * (TciC5 / (pciC5) ^ 0.5) + (k4 * (nC5 / 100)) * (TcnC5 / (pcnC5) ^ 0.5) + (k4 * (nC6 / 100)) * (TcnC6 / (pcnC6) ^ 0.5) + (k4 * (nC7 / 100)) * (TcnC7 / (pcnC7) ^ 0.5) + (k4 * (nC8 / 100)) * (TcnC8 / (pcnC8) ^ 0.5) + (k4 * (C2H4 / 100)) * (TcC2H4 / (pcC2H4) ^ 0.5) + (k4 * (C3H6 / 100)) * (TcC3H6 / (pcC3H6) ^ 0.5)
'Tpc: Pseudocritical Temperature
        Tpc = k * k / j
'Ppc: Pseudocritical Pressure
        Ppc = Tpc / j
'Tpr: Pseudoreduced Temperature
        Tpr = (T + 460) / Tpc
'GSG: Gas Specific Gravity Calculation
        GSG = ((H2S * MWH2S + CO2 * MWCO2 + n2 * MWN2 + c1 * MWC1 + c2 * MWC2 + C3 * MWC3 + iC4 * MWiC4 + nC4 * MWnC4 + (iC5 * MWiC5) + nC5 * MWnC5 + nC6 * MWnC6 + nC7 * MWnC7 + nC8 * MWnC8 + C2H4 * MWC2H4 + C3H6 * MWC3H6) / 100) / 29
        lnGSG = Log(GSG)
'Temperature T is in ºF, Ta is in kelvin and water vapor pressure is in mmHg (WVP, mmHg*0.0193 = WVP, psi)
        Ta = (T + 459.67) / 1.8
        WVP = (10 ^ (ap + bp / Ta + Cp * Log(Ta) / Log(10) + dp * Ta + ep * Ta ^ 2)) * 0.0193
'Temperature T is in ºF, Tb is in Kelvin and viscosity is in cp
        Tb = (T + 459.67) / 1.8
        WLV = 10 ^ (av + (bv / Tb) + (Cv * Tb) + (dv * (Tb ^ 2)))

        Phr = Exp(a0 + a1 * Log(Tpr) + a2 * ((NaCl / MWNaCl + KCl / MWKCl + CaCl2 / MWCaCl2) * Tpr) / (GSG ^ 2) + a3 * ((NaCl / MWNaCl + KCl / MWKCl + CaCl2 / MWCaCl2) * Tpr) + a4 * ((CH3OH / MWCH3OH + EG / MWEG + TEG / MWTEG + GL / MWGL) * (H2S + CO2 + n2)) / (Tpr ^ 2) + a5 * (CH3OH / MWCH3OH + EG / MWEG + TEG / MWTEG + GL / MWGL) / Tpr + a6 * ((WLV ^ 4) * (Tpr ^ 2) * WVP) + a7 * (WVP / Tpr) + a8 * (Tpr ^ 2) + a9 * ((100 - (NaCl + KCl + CaCl2)) * (Tpr ^ 2)) + a10 * ((lnGSG) * (Tpr)) + a11 * ((lnGSG) * (Log(Tpr)) * (WVP)) + a12 * ((lnGSG ^ 2) * (Tpr ^ 2)) + a13 * ((lnGSG) * (Tpr) * (WVP ^ 2)) + a14 * ((C3 + iC4) * (Tpr ^ 6)) + a15 * ((C3 + iC4) * (WVP) * (Tpr ^ 6)) + a16 * ((H2S + CO2) * (GSG) / Tpr) + a17 * ((n2) * (GSG) / (Tpr ^ 2)))
'ph: Hydrate-Formation Pressure
        Ph = (Phr * Ppc) - 14.7
        If HydrateAmeripourS > Ph Then HydrateAmeripourS = "": Exit Function
        HydrateAmeripourS = Ph
    Next x
    If HydrateAmeripourS > 12000 Then HydrateAmeripourS = "": Exit Function
    Exit Function
End If
HydrateAmeripourS = "Isnt clear if Pres or Temp is requested"
End Function
Private Function CheckIfThisIsTheRequestedComponent(Value As Variant, PossibleNames As Variant) As Boolean
    Value = UCase(Value)
    Dim NumberOfPossiblesNames As Integer
    Dim x As Integer
    CheckIfThisIsTheRequestedComponent = False
    NumberOfPossiblesNames = UBound(PossibleNames)
    For x = 1 To NumberOfPossiblesNames
        If Value = PossibleNames(x) Then CheckIfThisIsTheRequestedComponent = True: Exit For
    Next x
End Function
Function ReadAnotherSourceOfInformation(LabelsOfComponents As Range, ValuesOfComponents_MolPercentage As Range) As Variant
    Dim i As Integer
    Dim x As Byte
    Dim numberOfComponents As Integer
    Dim CheckForItemsNotRecognized() As Variant
    Dim Labels() As Variant
    Dim CO2 As Double, n2 As Double, c1 As Double, c2 As Double, C3 As Double, iC4 As Double, nC4 As Double, iC5 As Double, nC5 As Double, nC6 As Double, nC7 As Double, nC8 As Double, O2 As Double
    numberOfComponents = LabelsOfComponents.Columns.Count
'This part of the code will check if we are working with columns or rows
    If numberOfComponents = 1 Then
        numberOfComponents = LabelsOfComponents.Rows.Count
        ReDim Labels(1 To numberOfComponents)
        Labels = WorksheetFunction.Transpose(LabelsOfComponents.Value2)
        'I want to display an error if the size of the labels is different to the size of the values
        If numberOfComponents <> ValuesOfComponents_MolPercentage.Rows.Count Then ReadAnotherSourceOfInformation = "Problem in data Selection": Exit Function
    Else
        ReDim Labels(1 To numberOfComponents)
        For x = 1 To numberOfComponents
            Labels(x) = LabelsOfComponents(x)
        Next x
        'I want to display an error if the size of the labels is different to the size of the values
        If numberOfComponents <> ValuesOfComponents_MolPercentage.Columns.Count Then ReadAnotherSourceOfInformation = "Problem in data Selection": Exit Function
    End If
    ReDim CheckForItemsNotRecognized(1 To numberOfComponents)
    'Convert the ranges in variables.
    For i = 1 To numberOfComponents
        ReDim AUX(1 To 3)
        AUX(1) = "CO2"
        AUX(2) = "CARBON DIOXIDE"
        AUX(3) = "CARBONIC ACID"
        If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then CO2 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
        ReDim AUX(1 To 3)
        AUX(1) = "N2"
        AUX(2) = "NITROGEN"
        AUX(3) = "N"
        If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then n2 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
        ReDim AUX(1 To 3)
        AUX(1) = "C1"
        AUX(2) = "CH4"
        AUX(3) = "METHANE"
        If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then c1 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
        ReDim AUX(1 To 3)
        AUX(1) = "C2"
        AUX(2) = "C2H6"
        AUX(3) = "ETHANE"
        If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then c2 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
        ReDim AUX(1 To 3)
        AUX(1) = "C3"
        AUX(2) = "C3H8"
        AUX(3) = "PROPANE"
        If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then C3 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
        ReDim AUX(1 To 3)
        AUX(1) = "IC4"
        AUX(2) = "IC4H10"
        AUX(3) = "ISOBUTANE"
        If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then iC4 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
        ReDim AUX(1 To 6)
        AUX(1) = "NC4"
        AUX(2) = "C4"
        AUX(3) = "NC4H10"
        AUX(4) = "NORMALBUTANE"
        AUX(5) = "NORMAL BUTANE"
        AUX(6) = "BUTANE"
        If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then nC4 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
        ReDim AUX(1 To 3)
        AUX(1) = "IC5"
        AUX(2) = "IC5H12"
        AUX(3) = "ISOPENTANE"
        If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then iC5 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
        ReDim AUX(1 To 6)
        AUX(1) = "NC5"
        AUX(2) = "C5"
        AUX(3) = "NC5H12"
        AUX(4) = "NORMALPENTANE"
        AUX(5) = "NORMAL PENTANE"
        AUX(6) = "PENTANE"
        If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then nC5 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
        ReDim AUX(1 To 4)
        AUX(1) = "NC6"
        AUX(2) = "C6"
        AUX(3) = "C6H14"
        AUX(4) = "HEXANE"
        If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then nC6 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
        ReDim AUX(1 To 4)
        AUX(1) = "NC7"
        AUX(2) = "C7"
        AUX(3) = "C7H16"
        AUX(4) = "HEPTANE"
        If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then nC7 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
        ReDim AUX(1 To 5)
        AUX(1) = "NC8"
        AUX(2) = "C8"
        AUX(3) = "C8H18"
        AUX(4) = "C8H18+"
        AUX(5) = "OCTANE"
        If CheckIfThisIsTheRequestedComponent(Labels(i), AUX) Then nC8 = ValuesOfComponents_MolPercentage(i): CheckForItemsNotRecognized(i) = 1: GoTo NextIteration
NextIteration:
        Erase AUX
    Next i
'
'I want to display an error if a component is not recognized
    For i = 1 To numberOfComponents
        If CheckForItemsNotRecognized(i) <> 1 Then ReadAnotherSourceOfInformation = "ItemDontRecognized": Exit Function
    Next i
'
    Dim Answer(1 To 13)
    Answer(1) = n2
    Answer(2) = CO2
    Answer(3) = c1
    Answer(4) = c2
    Answer(5) = C3
    Answer(6) = iC4
    Answer(7) = nC4
    Answer(8) = iC5
    Answer(9) = nC5
    Answer(10) = nC6
    Answer(11) = nC7
    Answer(12) = nC8
    Answer(13) = O2
    ReadAnotherSourceOfInformation = WorksheetFunction.Transpose(Answer)
End Function
