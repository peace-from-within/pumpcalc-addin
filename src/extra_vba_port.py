import math

def CalcPumpEffAnother(DesRate, SpGr, DiffPress):
    Q = 4.4029 * DesRate
    H = (DiffPress * 10 / SpGr) * 3.28084
    if Q > 1000:
        return None
    if Q < 25:
        return None
    if Q < 100:
        RealQ = Q
        Q = 100
    else:
        RealQ = 200
    a = 80
    b = -0.2855 * H
    c = 0.000378 * Q * H
    D = -0.000000238 * Q ** 2 * H
    e = 0.000539 * H ** 2
    f = -0.000000639 * Q * H ** 2
    G = 0.0000000004 * Q ** 2 * H ** 2
    result = a + b + c + D + e + f + G
    if RealQ < 100:
        result = result - 0.35 * (100 - RealQ)
    return result

def SelMotorPower(DesRate, SpGr, DiffPress):
    DiffHead = DiffPress * 10 / SpGr
    GPM = 4.4029 * DesRate
    if GPM < 10:
        PumpEff = 6.9
    else:
        PumpEff = (-1.1237557 * math.log(GPM) + 20.3592672) * math.log(GPM) - 0.79087 - 76.536328 / math.log(GPM)
    HydPower = DesRate * DiffPress / 36.73
    BrkPower = HydPower / (PumpEff / 100)
    Y = BrkPower / 0.746
    MotEff = (-0.30546097 * math.log(Y) + 5.323494) * math.log(Y) + 72
    MotPower = BrkPower / (MotEff / 100)
    Z = MotPower
    steps = [0.4, 0.75, 1.5, 2.2, 3.7, 5.5, 7.5, 11, 15, 19, 22, 30, 37, 45, 55, 75, 110, 150]
    for s in steps:
        if Z <= s:
            return s
    return Z

def EqLenFactor(a, b):
    x = b
    Temp = a
    if not (Temp > 0):
        return None
    f = [0.0]*18
    f[1] = 0.00001*x**2 + 0.0275*x - 0.066
    f[2] = 0.00002*x**2 + 0.0448*x + 0.04
    f[3] = 0.00005*x**2 + 0.0602*x + 0.2518
    f[4] = 0.0001*x**2 + 0.0895*x + 0.5985
    f[5] = 0.0002*x**2 + 0.144*x + 0.6265
    f[6] = 0.0003*x**2 + 0.1901*x + 1.0593
    f[7] = 0.0005*x**2 + 0.2092*x + 1.8022
    f[8] = 0.0007*x**2 + 0.2659*x + 1.8858
    f[9] = 0.0013*x**2 + 0.3434*x + 2.7813
    f[10] = 0.0026*x**2 + 0.4148*x + 3.5215
    f[11] = 0.0038*x**2 + 0.507*x + 4.2357
    f[12] = 0.0039*x**2 + 0.72*x + 4.2978
    f[13] = 0.0099*x**2 + 0.8173*x + 4.5633
    f[14] = 0.0142*x**2 + 1.031*x + 5.1436
    f[15] = 0.0219*x**2 + 1.2168*x + 5.8647
    f[16] = 0.1437*x**2 + 2.8437*x + 10.228
    f[17] = 1.9048*x**2 + -3.8095*x + 44.286

    if Temp < f[1]:
        return 1.2
    elif Temp < f[2]:
        return 1.2 + (Temp-f[1])/(f[2]-f[1])*(1.5-1.2)
    elif Temp < f[3]:
        return 1.5 + (Temp-f[2])/(f[3]-f[2])*(1.7-1.5)
    elif Temp < f[4]:
        return 1.7 + (Temp-f[3])/(f[4]-f[3])*(2-1.7)
    elif Temp < f[5]:
        return 2 + (Temp-f[4])/(f[5]-f[4])*(2.25-2)
    elif Temp < f[6]:
        return 2.25 + (Temp-f[5])/(f[6]-f[5])*(2.5-2.25)
    elif Temp < f[7]:
        return 2.5 + (Temp-f[6])/(f[7]-f[6])*(2.75-2.5)
    elif Temp < f[8]:
        return 2.75 + (Temp-f[7])/(f[8]-f[7])*(3-2.75)
    elif Temp < f[9]:
        return 3 + (Temp-f[8])/(f[9]-f[8])*(3.25-3)
    elif Temp < f[10]:
        return 3.25 + (Temp-f[9])/(f[10]-f[9])*(3.5-3.25)
    elif Temp < f[11]:
        return 3.5 + (Temp-f[10])/(f[11]-f[10])*(3.75-3.5)
    elif Temp < f[12]:
        return 3.75 + (Temp-f[11])/(f[12]-f[11])*(4-3.75)
    elif Temp < f[13]:
        return 4 + (Temp-f[12])/(f[13]-f[12])*(4.25-4)
    elif Temp < f[14]:
        return 4.25 + (Temp-f[13])/(f[14]-f[13])*(4.5-4.25)
    elif Temp < f[15]:
        return 4.5 + (Temp-f[14])/(f[15]-f[14])*(4.75-4.5)
    elif Temp < f[16]:
        return 4.75 + (Temp-f[15])/(f[16]-f[15])*(7.25-4.75)
    elif Temp < f[17]:
        return 7.25 + (Temp-f[16])/(f[17]-f[16])*(14.75-7.25)
    else:
        return None  # "Out of range"
