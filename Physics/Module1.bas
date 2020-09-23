Attribute VB_Name = "Module1"
Public mode
Public Const PI As Double = "3,1415926535897932384626433832795"
Public Function ConvDegToRad(ByVal Deg As Double) As Double
    ConvDegToRad = (Deg / 180) * PI
End Function
Public Function ConvRadToDeg(value)
    ConvRadToDeg = value * (180 / PI)
End Function
