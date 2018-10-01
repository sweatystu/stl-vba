Option Explicit

' Custom errors
' Used to propogate and display errrors
Public custErr As New cCustomErrors

' Application Properties
' Used to optimise settings to speed up macros
' Must be initialised by adding app.Initialise at the beginning of a macro
Public app As New cAppProperties
