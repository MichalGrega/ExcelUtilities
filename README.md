# ExcelUtilities
Useful excel macros for all kinds of stuff.

This repository was created to keep me from writing the same code again and again.

# Type conversion

|Name|Description|Dependancies|
|:----|----|----|
|[fCLng](Functions/Strings/fClng.bas)|conversiont to Long type regardless of used decimal separator. Helps when there is an uncertainty about a decimal separator used.||
|[fConv](Functions/Numbers/fConv.bas)|conversion of decimal number into fixed length number of variable base. Can be also used as a counter with customizable scales. For more details check the function description in the module.|[fCLng](Functions/Strings/fClng.bas)|
|[DUMP](Functions/Utils/DUMP.bas)|function serializes (almost) any type of variable. It was originaly meant to serialize a dictionary, but because a dictionary item can be of any type, also this function needs to dump any type also including arrays reagardless of number of dimensions.|[fConv](Functions/Numbers/fConv.bas), [NumberOfArrayDimensions](Complete_Modules/modArraySupport.bas)|