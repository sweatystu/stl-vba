Option Explicit

' Custom errors
' Used to propogate and display errrors
Public custErr As New cCustomErrors

' Application Properties
' Used to optimise settings to speed up macros
' Must be initialised by adding app.Initialise at the beginning of a macro
Public app As New cAppProperties

' Progress User Form class
' Used to display progress of the macro
' Must be initialised by adding progress.LoadForm "Title" "Description" to the beginning of a macro
Public progress As New cProgress
