
								'Declaring a variables at the top acts a placeholder memory for your variables
Option Explcit 					'having this statement at the top requires to declare your variables
Dim FinalRow As Integer 		'Dim or Private are similar. The scope of the variables is only within a subroutine
Public FinalRow2 As Integer 	'This scope will allow to use the variables for every single sub-routine within a module

			   
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~			
'The Static declaration and Variable Expiry 
'STATIC means that the value will be in memory and each time you run your macro 
'the value will increase by 1 as shown below
'so the value is retained in memory 

Sub StaticExample ()

Static y as Long

y = y+1
msgbox y

End Sub


' To reset the value of y you can create a small sub-routine

Sub resetVariable ()
 End
End Sub	

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Constants for variables
'You can also create constants which later can be used in every single sub-routine within a module. Again
'the constants declaration has to be done at the very top under "Option Explicit" and you won't have to hardcode the 
'value for the constant value

Option Explicit 
Public Const power3 As Integer = 3

Sub ConstantTest ()
Dim x As Integer
x = 2
result = x^power3
MsgBox result

End Sub	

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'ARRAYS is a group of variables that all have the same data type and the same name

'

