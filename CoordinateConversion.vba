'
'   Algorithm: http://zoologie.umons.ac.be/tc/algorithms.aspx
'   Implemented by: Teun Reyniers
'   The code is distributed as: GNU General Public License
'

Const PI As Double = 3.14159265358979
Type llCoorinate
    lat As Double
    lng As Double
End Type

Type xyCoorinate
    X As Double
    Y As Double
End Type

'visible excel formulas

Public Function BLam72ToBDatum_Lat(ByRef X As Double, Y As Double) As Double

    BLam72ToSpere_Lat = BelgianLambert72ToSpherical(X, Y).lat
    
End Function

Public Function BLam72ToBDatum_Lng(ByRef X As Double, Y As Double) As Double

    BLam72ToSpere_Lng = BelgianLambert72ToSpherical(X, Y).lng
    
End Function

Public Function SpereToBLam72_X(ByRef lat As Double, lng As Double) As Double

    SpereToBLam72_X = SphericalToBelgianLambert72(lat, lng).X
    
End Function

Public Function SpereToBLam72_Y(ByRef lat As Double, lng As Double) As Double

    SpereToBLam72_Y = SphericalToBelgianLambert72(lat, lng).Y
    
End Function

Public Function BDatumToWGS84_Lat(ByRef lat As Double, ByRef lng As Double, Optional ByRef haut As Double = 0) As Double

    BDatumToWGS84_Lat = BelgianDatumToWGS84(lat, lng, haut).lat

End Function

Public Function BDatumToWGS84_Lng(ByRef lat As Double, ByRef lng As Double, Optional ByRef haut As Double = 0) As Double

    BDatumToWGS84_Lng = BelgianDatumToWGS84(lat, lng, haut).lng

End Function

Public Function WGS84ToBDatum_Lat(ByRef lat As Double, ByRef lng As Double, Optional ByRef haut As Double = 0) As Double

    WGS84ToBDatum_Lat = WGS84ToBelgianDatum(lat, lng, haut).lat

End Function

Public Function WGS84ToBDatum_Lng(ByRef lat As Double, ByRef lng As Double, Optional ByRef haut As Double = 0) As Double

    WGS84ToBDatum_Lng = WGS84ToBelgianDatum(lat, lng, haut).lng

End Function

Public Function BLam72ToWGS84_Lat(ByRef X As Double, Y As Double, Optional ByRef haut As Double = 0) As Double

    Dim sphere As llCoorinate
    sphere = BelgianLambert72ToSpherical(X, Y)
    BLam72ToWGS84_Lat = BelgianDatumToWGS84(sphere.lat, sphere.lng, haut).lat

End Function

Public Function BLam72ToWGS84_Lng(ByRef X As Double, Y As Double, Optional ByRef haut As Double = 0) As Double

    Dim sphere As llCoorinate
    sphere = BelgianLambert72ToSpherical(X, Y)
    BLam72ToWGS84_Lng = BelgianDatumToWGS84(sphere.lat, sphere.lng, haut).lng

End Function

Public Function WGS84ToBLam72_X(ByRef lat As Double, ByRef lng As Double, Optional ByRef haut As Double = 0) As Double

    Dim sphere As llCoorinate
    sphere = WGS84ToBelgianDatum(lat, lng, haut)
    WGS84ToBLam72_X = SphericalToBelgianLambert72(sphere.lat, sphere.lng).X

End Function

Public Function WGS84ToBLam72_Y(ByRef lat As Double, ByRef lng As Double, Optional ByRef haut As Double = 0) As Double

    Dim sphere As llCoorinate
    sphere = WGS84ToBelgianDatum(lat, lng, haut)
    WGS84ToBLam72_Y = SphericalToBelgianLambert72(sphere.lat, sphere.lng).Y

End Function

'convertion
Private Function BelgianLambert72ToSpherical(X As Double, Y As Double) As llCoorinate
    
    ' Belgian Lambert 72 to spherical coordinates
    '
    ' Belgian Lambert 1972---> Spherical coordinates
    ' Input parameters : X, Y = Belgian coordinates in meters
    ' Output : latitude and longitude in Belgium Datum!
    '
    
    Const LongRef As Double = 0.076042943        '=4°21'24"983
    Const nLamb As Double = 0.7716421928
    Const aCarre As Double = 6378388 ^ 2
    Const bLamb As Double = 6378388 * (1 - (1 / 297))
    Const eCarre As Double = (aCarre - bLamb ^ 2) / aCarre
    Const KLamb As Double = 11565915.812935
     
    Dim eLamb As Double
    eLamb = Sqr(eCarre)
    Dim eSur2 As Double
    eSur2 = eLamb / 2
     
    Dim Tan1 As Double
    Tan1 = (X - 150000.01256) / (5400088.4378 - Y)
    Dim Lambda As Double
    Lambda = LongRef + (1 / nLamb) * (0.000142043 + Atn(Tan1))
    Dim RLamb As Double
    RLamb = Sqr((X - 150000.01256) ^ 2 + (5400088.4378 - Y) ^ 2)
     
    Dim TanZDemi As Double
    TanZDemi = (RLamb / KLamb) ^ (1 / nLamb)
    Dim Lati1 As Double
    Latil = 2 * Atn(TanZDemi)
     
    Dim eSin As Double
    Dim Mult1, Mult2, Mult As Double
    Dim LatiN, Diff As Double
     
    Dim lat, lng As Double
     
    Do
        eSin = eLamb * Sin(Lati1)
        Mult1 = 1 - eSin
        Mult2 = 1 + eSin
        Mult = (Mult1 / Mult2) ^ (eLamb / 2)
        LatiN = (PI / 2) - (2 * (Atn(TanZDemi * Mult)))
        Diff = LatiN - Lati1
        Lati1 = LatiN
    Loop While Math.Abs(Diff) > 0.0000000277777
    
    BelgianLambert72ToSpherical.lat = (LatiN * 180) / PI
    BelgianLambert72ToSpherical.lng = (Lambda * 180) / PI

End Function

Private Function SphericalToBelgianLambert72(lat As Double, lng As Double) As xyCoorinate

    ' Spherical coordinates to Belgian Lambert 72
    '
    ' Conversion from spherical coordinates to Lambert 72
    ' Input parameters : lat, lng (spherical coordinates)
    ' Spherical coordinates are in decimal degrees converted to Belgium datum!
    '
    
    Const LongRef As Double = 0.076042943        '=4°21'24"983
    Const bLamb As Double = 6378388 * (1 - (1 / 297))
    Const aCarre As Double = 6378388 ^ 2
    Const eCarre As Double = (aCarre - bLamb ^ 2) / aCarre
    Const KLamb As Double = 11565915.812935
    Const nLamb As Double = 0.7716421928
    
    Dim eLamb As Double
    eLamb = Sqr(eCarre)
    Dim eSur2 As Double
    eSur2 = eLamb / 2
    
    'conversion to radians
    lat = (PI / 180) * lat
    lng = (PI / 180) * lng
    
    Dim eSinLatitude As Double
    eSinLatitude = eLamb * Sin(lat)
    Dim TanZDemi As Double
    TanZDemi = (Tan((PI / 4) - (lat / 2))) * (((1 + (eSinLatitude)) / (1 - (eSinLatitude))) ^ (eSur2))
    
    Dim RLamb As Double
    RLamb = KLamb * ((TanZDemi) ^ nLamb)
    
    Dim Teta As Double
    Teta = nLamb * (lng - LongRef)
    
    SphericalToBelgianLambert72.X = 150000 + 0.01256 + RLamb * Sin(Teta - 0.000142043)
    SphericalToBelgianLambert72.Y = 5400000 + 88.4378 - RLamb * Cos(Teta - 0.000142043)
    
End Function

Private Function BelgianDatumToWGS84(lat As Double, lng As Double, haut As Double) As llCoorinate

    ' Belgian Datum to WGS84 conversion (Molodensky 3 parameters)
    '
    ' Input parameters : Lat, Lng : latitude / longitude in decimal degrees and in Belgian 1972 datum
    ' Output parameters : LatWGS84, LngWGS84 : latitude / longitude in decimal degrees and in WGS84 datum
    '
     
    'Const Haut = 0      'Altitude
    Dim LatWGS84, LngWGS84 As Double
    Dim DLat, DLng As Double
    Dim Dh As Double
    Dim dy, dx, dz As Double
    Dim da, df As Double
    Dim LWa, Rm, Rn, LWb As Double
    Dim LWf, LWe2 As Double
    Dim SinLat, SinLng As Double
    Dim CoSinLat As Double
    Dim CoSinLng As Double
     
    Dim Adb As Double
     
    'conversion to radians
    lat = (PI / 180) * lat
    lng = (PI / 180) * lng
     
    SinLat = Sin(lat)
    SinLng = Sin(lng)
    CoSinLat = Cos(lat)
    CoSinLng = Cos(lng)
     
    dx = -125.8
    dy = 79.9
    dz = -100.5
    da = -251#
    df = -0.000014192702
     
    LWf = 1 / 297
    LWa = 6378388
    LWb = (1 - LWf) * LWa
    LWe2 = (2 * LWf) - (LWf * LWf)
    Adb = 1 / (1 - LWf)
     
    Rn = LWa / Sqr(1 - LWe2 * SinLat * SinLat)
    Rm = LWa * (1 - LWe2) / (1 - LWe2 * lat * lat) ^ 1.5
     
    DLat = -dx * SinLat * CoSinLng - dy * SinLat * SinLng + dz * CoSinLat
    DLat = DLat + da * (Rn * LWe2 * SinLat * CoSinLat) / LWa
    DLat = DLat + df * (Rm * Adb + Rn / Adb) * SinLat * CoSinLat
    DLat = DLat / (Rm + haut)
     
    DLng = (-dx * SinLng + dy * CoSinLng) / ((Rn + haut) * CoSinLat)
    Dh = dx * CoSinLat * CoSinLng + dy * CoSinLat * SinLng + dz * SinLat
    Dh = Dh - da * LWa / Rn + df * Rn * lat * lat / Adb
     
    BelgianDatumToWGS84.lat = ((lat + DLat) * 180) / PI
    BelgianDatumToWGS84.lng = ((lng + DLng) * 180) / PI
    
End Function

Private Function WGS84ToBelgianDatum(lat As Double, lng As Double, haut As Double) As llCoorinate
    '
    'WGS84 to Belgian Datum conversion (Molodensky 3 parameters)
    '
    'Input parameters : Lat, Lng : latitude / longitude in decimal degrees and in WGS84 datum
    'Output parameters : LatBel, LngBel : latitude / longitude in decimal degrees and in Belgian datum
    '
     
    'Const Haut = 0      'Altitude
    Dim LatBel, LngBel As Double
    Dim DLat, DLng As Double
    Dim Dh As Double
    Dim dy, dx, dz As Double
    Dim da, df As Double
    Dim LWa, Rm, Rn, LWb As Double
    Dim LWf, LWe2 As Double
    Dim SinLat, SinLng As Double
    Dim CoSinLat As Double
    Dim CoSinLng As Double
     
    Dim Adb As Double
     
    'conversion to radians
    lat = (PI / 180) * lat
    lng = (PI / 180) * lng
     
    SinLat = Sin(lat)
    SinLng = Sin(lng)
    CoSinLat = Cos(lat)
    CoSinLng = Cos(lng)
     
    dx = 125.8
    dy = -79.9
    dz = 100.5
    da = 251#
    df = 0.000014192702
     
    LWf = 1 / 297
    LWa = 6378388
    LWb = (1 - LWf) * LWa
    LWe2 = (2 * LWf) - (LWf * LWf)
    Adb = 1 / (1 - LWf)
     
    Rn = LWa / Sqr(1 - LWe2 * SinLat * SinLat)
    Rm = LWa * (1 - LWe2) / (1 - LWe2 * lat * lat) ^ 1.5
     
    DLat = -dx * SinLat * CoSinLng - dy * SinLat * SinLng + dz * CoSinLat
    DLat = DLat + da * (Rn * LWe2 * SinLat * CoSinLat) / LWa
    DLat = DLat + df * (Rm * Adb + Rn / Adb) * SinLat * CoSinLat
    DLat = DLat / (Rm + haut)
     
    DLng = (-dx * SinLng + dy * CoSinLng) / ((Rn + haut) * CoSinLat)
    Dh = dx * CoSinLat * CoSinLng + dy * CoSinLat * SinLng + dz * SinLat
    Dh = Dh - da * LWa / Rn + df * Rn * lat * lat / Adb
     
    WGS84ToBelgianDatum.lat = ((lat + DLat) * 180) / PI
    WGS84ToBelgianDatum.lng = ((lng + DLng) * 180) / PI
    
End Function
