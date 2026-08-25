import math

def Log10(x):
    return math.log(x) / math.log(10.0)

# ---- Calc_J / Calc_JA / Calc_RA ----
def Calc_J(ValU, RAAss, ValAS, ValAP):
    if ValU == 1:
        return RAAss * ValAS * ValAP
    else:
        return ValAS * ValAP + 2 * ValAS * math.log(RAAss) + 1 - RAAss ** 2

def Calc_JA(ValU, RAAss, ValAS):
    if ValU == 2:
        return 2 * ValAS / RAAss - 2 * RAAss
    else:
        return 0

def Calc_RA(ValU, ValAS, ValAP, max_iter=100000):
    diff = 1000
    oldRA = 1
    n = 0
    while diff > 0.0000000001:
        ValJ = Calc_J(ValU, oldRA, ValAS, ValAP)
        ValJA = Calc_JA(ValU, oldRA, ValAS)
        if ValU == 1:
            newRA = ValJ + ValJ * math.log(oldRA) / ValAP + 1
        else:
            newRA = -1 * ValJ / ValJA + oldRA
        diff = abs(newRA - oldRA)
        oldRA = newRA
        n += 1
        if n > max_iter:
            raise RuntimeError("Calc_RA did not converge")
    return oldRA, n

# ---- Pipe_ID (needs the Pipe_ID sheet grid) ----
def Pipe_ID(norm_id, sch, cellget, max_row=200, max_col=60):
    # cellget(row, col) -> raw cell value (None for blank), 1-indexed like VBA Cells(row,col)
    def v0(x):
        return 0 if x is None else x
    int_I = 4
    NID = 0
    # pre-test "Do Until Norm_ID = NID ... Loop"
    while not (norm_id == v0(NID)):
        int_I += 1
        if int_I > max_row:
            raise RuntimeError("Pipe_ID row search exceeded sheet")
        NID = cellget(int_I, 1)
    int_X = int_I

    int_J = 1
    Sch_X = 0
    while not (sch == v0(Sch_X)):
        int_J += 1
        if int_J > max_col:
            raise RuntimeError("Pipe_ID col search exceeded sheet")
        Sch_X = cellget(3, int_J)
    int_Y = int_J

    return cellget(int_X, int_Y)

# ---- Two-phase basic props ----
def _common(WtLiqM, WtVapM, DenLiqM, DenVapM):
    Wliq = WtLiqM * 2.2046
    Wvap = WtVapM * 2.2046
    Wtot = Wliq + Wvap
    DenLiq = DenLiqM * 0.0624
    DenVap = DenVapM * 0.0624
    DenHomo = Wtot / (Wliq / DenLiq + Wvap / DenVap)
    return Wliq, Wvap, Wtot, DenLiq, DenVap, DenHomo

def VaporVolFrac(WtLiqM, WtVapM, DenLiqM, DenVapM):
    Wliq, Wvap, Wtot, DenLiq, DenVap, DenHomo = _common(WtLiqM, WtVapM, DenLiqM, DenVapM)
    QL = Wliq / DenLiq / 3600
    QV = Wvap / DenVap / 3600
    return QV / (QV + QL)

def AvgDensity(WtLiqM, WtVapM, DenLiqM, DenVapM):
    Wliq, Wvap, Wtot, DenLiq, DenVap, DenHomo = _common(WtLiqM, WtVapM, DenLiqM, DenVapM)
    return DenHomo / 0.0624

def DPHomogeneous(WtLiqM, WtVapM, DenLiqM, DenVapM, e, PipeID):
    Wliq, Wvap, Wtot, DenLiq, DenVap, DenHomo = _common(WtLiqM, WtVapM, DenLiqM, DenVapM)
    ID = PipeID / 12
    Re = 6.316 * Wtot / (ID * 12) / 0.1
    a = -2 * Log10((e / ID / 3.7 + 12 / Re))
    b = -2 * Log10((e / ID / 3.7 + 2.51 * a / Re))
    c = -2 * Log10((e / ID / 3.7 + 2.51 * b / Re))
    f = (a - (b - a) ** 2 / (c - 2 * b + a)) ** (-2)
    return (f * 0.000336 * Wtot ** 2 / DenHomo / (ID * 12) ** 5) / 4.42075824

def DPDukler(WtLiqM, WtVapM, DenLiqM, DenVapM, VisLiq, VisVap, PipeID, e, max_iter=100000):
    Wliq, Wvap, Wtot, DenLiq, DenVap, DenHomo = _common(WtLiqM, WtVapM, DenLiqM, DenVapM)
    Diam = PipeID / 12
    AREA = Diam ** 2 * 3.14 / 4
    QL = Wliq / DenLiq / 3600
    QV = Wvap / DenVap / 3600
    F_v = QV / (QV + QL)
    F_l = 1 - F_v
    Vm = (QV + QL) / AREA
    NFR = Vm ** 2 / 32.174 / Diam
    x = Wvap / (Wliq + Wvap)
    GT = Wtot / AREA

    RvCal = 0.5
    RvAssum = None
    n1 = 0
    while RvAssum is None or abs(RvAssum - RvCal) >= 0.0000000001:
        RvAssum = RvCal
        NRe1 = Diam * GT / (2.42 * (RvAssum * VisVap + (1 - RvAssum) * VisLiq))
        Z = NRe1 ** (1 / 6) * NFR ** (1 / 8) / F_l ** 0.25
        if Z < 10:
            Charlie = -0.16367 + 0.31037 * Z - 0.03525 * Z ** 2 + 0.001366 * Z ** 3
        else:
            Charlie = 0.75545 + 0.003585 * Z - 0.1436 * 10 ** (-4) * Z ** 2
        RvCal = Charlie / ((1 / x - 1) * (DenVap / DenLiq) + 1)
        n1 += 1
        if n1 > max_iter:
            raise RuntimeError("Rv loop no converge")
    RV = RvCal

    DENIP1 = DenVap * RV + DenLiq * (1 - RV)
    LF_l = math.log(F_l)
    H = 1 - LF_l / (1.281 + 0.478 * LF_l + 0.444 * LF_l ** 2 + 0.094 * LF_l ** 3 + 0.00843 * LF_l ** 4)
    DenNS = DenLiq * F_l + DenVap * (1 - F_l)
    VisNS = (VisLiq * F_l + VisVap * (1 - F_l)) * 2.42
    Bravo1 = (DenLiq / DenNS) * (F_l ** 2 / (1 - RV)) + (DenVap / DenNS) * ((1 - F_l) ** 2 / RV)
    ReTP1 = 4 * Wtot * Bravo1 / (3.14 * Diam * VisNS)

    fcal = 0.001
    fassum = None
    n2 = 0
    while fassum is None or abs(fcal - fassum) >= 0.00000001:
        fassum = fcal
        fcal = 1 / (-2 * Log10(e / Diam / 3.7 + (2.51 / ReTP1 / fassum))) ** 2
        n2 += 1
        if n2 > max_iter:
            raise RuntimeError("f loop no converge")
    f = fcal / 4

    DPDukler_val = 2 * (GT / 3600) ** 2 * f * 100 * Bravo1 * H / (32.2 * Diam * DenNS * 144) / 4.42075824
    return DPDukler_val, n1, n2

def InplaceDensity(WtLiqM, WtVapM, DenLiqM, DenVapM, VisLiq, VisVap, PipeID, max_iter=100000):
    Wliq = WtLiqM * 2.2046
    Wvap = WtVapM * 2.2046
    Wtot = Wliq + Wvap
    DenLiq = DenLiqM * 0.0624
    DenVap = DenVapM * 0.0624
    Diam = PipeID / 12
    AREA = Diam ** 2 * 3.14 / 4
    QL = Wliq / DenLiq / 3600
    QV = Wvap / DenVap / 3600
    F_v = QV / (QV + QL)
    Fl = 1 - F_v
    Vm = (QV + QL) / AREA
    NFR = Vm ** 2 / 32.174 / Diam
    x = Wvap / (Wliq + Wvap)
    GT = Wtot / AREA

    RvCal = 0.5
    RvAssum = None
    n = 0
    while RvAssum is None or abs(RvAssum - RvCal) >= 0.0000000001:
        RvAssum = RvCal
        NRe1 = Diam * GT / (2.42 * (RvAssum * VisVap + (1 - RvAssum) * VisLiq))
        Z = NRe1 ** (1 / 6) * NFR ** (1 / 8) / Fl ** 0.25
        if Z < 10:
            Charlie = -0.16367 + 0.31037 * Z - 0.03525 * Z ** 2 + 0.001366 * Z ** 3
        else:
            Charlie = 0.75545 + 0.003585 * Z - 0.1436 * 10 ** (-4) * Z ** 2
        RvCal = Charlie / ((1 / x - 1) * (DenVap / DenLiq) + 1)
        n += 1
        if n > max_iter:
            raise RuntimeError("no converge")
    RV = RvCal
    return (DenVap * RV + DenLiq * (1 - RV)) / 0.0624, n

def TwoPhaseVelocity(WtLiqM, WtVapM, DenLiqM, DenVapM, PipeID):
    Wliq = WtLiqM * 2.2046
    Wvap = WtVapM * 2.2046
    DenLiq = DenLiqM * 0.0624
    DenVap = DenVapM * 0.0624
    Diam = PipeID / 12
    AREA = Diam ** 2 * 3.14 / 4
    QL = Wliq / DenLiq / 3600
    QV = Wvap / DenVap / 3600
    return (QV + QL) / AREA * 0.3048

def BakerXval(WtLiqM, WtVapM, DenLiqM, DenVapM, MWl, SurTen, VisLiq, PipeID):
    Wliq = WtLiqM * 2.2046
    Wvap = WtVapM * 2.2046
    DenLiq = DenLiqM * 0.0624
    DenVap = DenVapM * 0.0624
    Diam = PipeID / 12
    AREA = Diam ** 2 * 3.14 / 4
    if SurTen == 0:
        ST = (((57 + 2.3 * MWl) / MWl) * ((DenLiq - DenVap) / 62.3)) ** 4
    else:
        ST = SurTen
    La = ((DenVap / 0.075) * (DenLiq / 62.3)) ** 0.5
    P = 73 / ST * (VisLiq * (62.3 / DenLiq) ** 2) ** (1 / 3)
    G = Wvap / AREA
    L = Wliq / AREA
    return L * P * La / G

def BakerYval(WtVapM, DenLiqM, DenVapM, PipeID):
    Wvap = WtVapM * 2.2046
    DenLiq = DenLiqM * 0.0624
    DenVap = DenVapM * 0.0624
    Diam = PipeID / 12
    AREA = Diam ** 2 * 3.14 / 4
    La = ((DenVap / 0.075) * (DenLiq / 62.3)) ** 0.5
    G = Wvap / AREA
    return G / La

def Baker(X1, Y1):
    if X1 < 0.1 or X1 > 10000.0 or Y1 < 100.0 or Y1 > 100000.0:
        return "Out of Range"
    x = (Log10(X1) - Log10(0.1)) / (Log10(10000.0) - Log10(0.1))
    Y = (Log10(Y1) - Log10(100)) / (Log10(100000.0) - Log10(100))
    Line1 = 0.19015 * x ** 3 - 0.38395 * x ** 2 - 0.91636 * x + 0.94269
    Line2 = 0.41754 * x ** 3 - 0.99032 * x ** 2 + 0.04973 * x + 0.61096
    line3 = -3.59472 * x ** 3 + 5.82745 * x ** 2 - 3.13262 * x + 1.37817
    line4 = 7.9972 * x ** 3 - 7.0625 * x ** 2 + 1.35609 * x + 0.65341
    line5 = -0.36 * x + 0.535
    line6 = -133.333 * x ** 2 + 173 * x - 55.10667
    line7 = -9.18737 * x ** 3 + 23.70631 * x ** 2 - 21.7587 * x + 7.33735
    if x <= 0.294142932210028:
        if Y <= Line2: return "Stratified"
        elif Y <= Line1: return "Wave"
        elif Y <= line3: return "Annular"
        else: return "Dispersed"
    elif x <= 0.448315796039707:
        if Y <= Line2: return "Stratified"
        elif Y <= Line1: return "Wave"
        elif Y <= line4: return "Slug"
        elif Y <= line3: return "Annular"
        else: return "Dispersed"
    elif x <= 0.571278202693418:
        if Y <= Line1: return "Stratified"
        elif Y <= line4: return "Slug"
        elif Y <= line3: return "Annular"
        else: return "Dispersed"
    elif x <= 0.609274283833812:
        if Y <= Line1: return "Stratified"
        elif Y <= line5: return "Plug"
        elif Y <= line4: return "Slug"
        elif Y <= line3: return "Annular"
        else: return "Dispersed"
    elif x <= 0.637634634557965:
        if Y <= Line1: return "Stratified"
        elif Y <= line5: return "Plug"
        elif Y <= line4: return "Slug"
        elif Y <= line7: return "Annular"
        elif Y <= line6: return "Bubble or Froth"
        else: return "Dispersed"
    elif x <= 0.943008707285949:
        if Y <= Line1: return "Stratified"
        elif Y <= line5: return "Plug"
        elif Y <= line7: return "Slug"
        else: return "Bubble or Froth"
    else:
        if Y <= line7: return "Plug"
        else: return "Bubble or Froth"

def GriffithWallisXval(WtLiqM, WtVapM, DenLiqM, DenVapM, PipeID):
    Wliq = WtLiqM * 2.2046
    Wvap = WtVapM * 2.2046
    DenLiq = DenLiqM * 0.0624
    DenVap = DenVapM * 0.0624
    Diam = PipeID / 12
    AREA = Diam ** 2 * 3.14 / 4
    QL = Wliq / DenLiq / 3600
    QV = Wvap / DenVap / 3600
    Vm = (QV + QL) / AREA
    return Vm ** 2 / 32.174 / Diam

def GriffithWallisYval(WtLiqM, WtVapM, DenLiqM, DenVapM, PipeID):
    Wliq = WtLiqM * 2.2046
    Wvap = WtVapM * 2.2046
    DenLiq = DenLiqM * 0.0624
    DenVap = DenVapM * 0.0624
    QL = Wliq / DenLiq / 3600
    QV = Wvap / DenVap / 3600
    return QV / (QV + QL)

def Griffith(x, Y):
    if x < 0 or x > 400 or Y < 0 or Y > 1:
        return "Out of Range"
    Line1 = 0.06903 * x ** (-0.650874)
    Line2 = -0.00037612839 * x + 0.20009027081
    line3 = -0.0000034677 * x ** 3 + 0.0003366929 * x ** 2 - 0.0173144767 * x + 1.1758057232
    if x <= 0.24:
        return "Bubble" if Y <= Line1 else "Slug"
    elif x <= 80:
        if Y <= Line2: return "Bubble"
        elif Y <= line3: return "Slug"
        else: return "Annular Mist"
    else:
        return "Annular Mist"

def CalcPres(Head, SG):
    return Head * SG / 10

def CalcHead(Press, SG):
    return Press * 10 / SG

def EstNPSHR(MetricQ):
    Q = MetricQ * 4.403
    val = (Q ** 0.67 * 10 ** (-0.612)) / 3.281 + 0.6
    if val < 1.83:
        val = 1.83
    return val

def CalcPumpEff(FlowRate):
    GPM = 4.4029 * FlowRate
    if GPM < 10:
        return 6.9
    else:
        return (-1.1237557 * math.log(GPM) + 20.3592672) * math.log(GPM) - 0.79087 - 76.536328 / math.log(GPM)


def _max_vba(a, b, c):
    result = a
    if b > result:
        result = b
    elif c > result:
        result = c
    return result

def ConvCVPressDrop15(TotalSucFixedPress, TotalSucVarLossRtd, TotalDischFixedPress, TotalDischVarLossRtd, OverDesign):
    TotalFixedPress = TotalDischFixedPress - TotalSucFixedPress
    TotalVarLossRtd = TotalSucVarLossRtd + TotalDischVarLossRtd
    TotalLossRtd = TotalFixedPress + TotalVarLossRtd
    PercentVarLoss1 = TotalVarLossRtd / TotalLossRtd * 100
    x = OverDesign / 100
    if PercentVarLoss1 <= 25:
        Y = 10
    elif PercentVarLoss1 <= 40:
        Y1 = 10
        Y2 = -0.0343 * x + 11.714
        Y = Y1 + (PercentVarLoss1 - 25) / (40 - 25) * (Y2 - Y1)
    elif PercentVarLoss1 <= 50:
        Y1 = -0.0343 * x + 11.714
        Y2 = -0.0629 * x + 13.143
        Y = Y1 + (PercentVarLoss1 - 40) / (50 - 40) * (Y2 - Y1)
    elif PercentVarLoss1 <= 75:
        Y1 = -0.0629 * x + 13.143
        Y2 = -0.108 * x + 15.429
        Y = Y1 + (PercentVarLoss1 - 50) / (75 - 50) * (Y2 - Y1)
    elif PercentVarLoss1 <= 100:
        Y1 = -0.108 * x + 15.429
        Y2 = -0.149 * x + 17.429
        Y = Y1 + (PercentVarLoss1 - 75) / (100 - 75) * (Y2 - Y1)
    else:
        Y = -0.149 * x + 17.429
    return Y * TotalLossRtd / 100

def CVPressDrop(Method, CVType, FlagReflux, TotalSucFixedPress, TotalSucVarLossNor, TotalSucVarLossRtd,
                 TotalDischFixedPress, TotalDischVarLossNor, TotalDischVarLossRtd, OverDesign, M_Input):
    TotalFixedPress = TotalDischFixedPress - TotalSucFixedPress
    TotalVarLossNor = TotalSucVarLossNor + TotalDischVarLossNor
    TotalVarLossRtd = TotalSucVarLossRtd + TotalDischVarLossRtd
    TotalLossNor = TotalFixedPress + TotalVarLossNor
    TotalLossRtd = TotalFixedPress + TotalVarLossRtd
    SucPressNor = TotalSucFixedPress - TotalSucVarLossNor
    SucPressRtd = TotalSucFixedPress - TotalSucVarLossRtd

    if str(Method) == "No Control Valve":
        return 0
    elif str(Method) == "GTP Std. Method":
        if str(CVType) == "single":
            BasePressDrop = 0.77
        elif str(CVType) == "double":
            BasePressDrop = 0.49
        elif str(CVType) == "cage":
            BasePressDrop = 0.28
        elif str(CVType) == "butterfly":
            BasePressDrop = 0.01
        elif str(CVType) == "vball":
            BasePressDrop = 0.07
        else:
            BasePressDrop = 0.77
        DischPress = 1.1 * (((100 + OverDesign) / 100) ** 2 - 1) * TotalDischVarLossNor
        DischPress = DischPress + TotalDischFixedPress + TotalDischVarLossNor
        DischPress = 100 / 95 * (DischPress + BasePressDrop)
        return DischPress - TotalDischFixedPress - TotalDischVarLossNor
    elif str(Method) == "Lummus Method":
        if OverDesign <= 15:
            aaa = TotalVarLossNor * 0.5
            bbb = 1.055
            ccc = 0.08 * (SucPressNor + TotalLossNor)
            return _max_vba(aaa, bbb, ccc)
        else:  # OverDesign > 15
            zzz = ConvCVPressDrop15(TotalSucFixedPress, TotalSucVarLossRtd, TotalDischFixedPress, TotalDischVarLossRtd, OverDesign)
            aaa = 0.703
            if FlagReflux == "Reflux":
                CVPressDropRtd = _max_vba(zzz, aaa, 0)
            else:
                bbb = 0.08 * (SucPressRtd + TotalLossRtd)
                CVPressDropRtd = _max_vba(zzz, aaa, bbb)
            return CVPressDropRtd + TotalVarLossRtd - TotalVarLossNor
    elif str(Method) == "NODCO":
        CVPressDrop_val = TotalLossNor * 15 / 85
        CVPressDropRtd = CVPressDrop_val + TotalVarLossNor - TotalVarLossRtd
        if CVPressDropRtd < 0.7:
            CVPressDrop_val = 0.7 - TotalVarLossNor + TotalVarLossRtd
        return CVPressDrop_val
    else:
        return M_Input
